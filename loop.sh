#!/bin/bash

# Cascade Industries Test Suite - Ralph Loop Runner
# Usage:
#   ./loop.sh                    - Run build loop (build only)
#   ./loop.sh plan               - Run planning loop
#   ./loop.sh N                  - Run build loop for max N iterations
#   ./loop.sh plan N             - Run planning loop for max N iterations
#   ./loop.sh include-tests      - Run build loop with generator + grader checks
#   ./loop.sh include-tests N    - Same, capped at N iterations
#
# Model:
#   Build/plan iterations use Opus 4.6 (--model opus) for complex reasoning.
#
# By default, only the build phase runs. The optional "include-tests" phase
# runs pytest, ruff, and a determinism smoke test (generate twice, diff outputs)
# once generate_test_suite.py exists.
#
# FIX for Claude Code hang bug (GitHub #19060, #25629, #31050):
# Claude completes work but never calls process.exit(). The process hangs
# indefinitely at 0% CPU with stdout open. Using --output-format stream-json
# lets us detect the {"type":"result"} event and kill the process ourselves.

set -e
MODE="build"
INCLUDE_TESTS=false
MAX_ITERATIONS=0
ITERATION=0
BUILD_MODEL="opus"
HARD_TIMEOUT=2700  # 45min safety net (should never hit with stream-json detection)

# Project directory = directory containing this script
PROJECT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
TEMP_OUTPUT=$(mktemp)
trap "rm -f $TEMP_OUTPUT" EXIT

# Kill any orphaned Claude processes from previous runs
cleanup_orphan_claude_processes() {
    local current_ppid=$$
    ps aux | grep -E "claude.*-p.*--dangerously-skip-permissions" | grep -v grep | while read -r line; do
        local pid=$(echo "$line" | awk '{print $2}')
        if [ "$pid" != "$current_ppid" ]; then
            kill "$pid" 2>/dev/null || true
        fi
    done
}
cleanup_orphan_claude_processes

# Run claude with stream-json and detect completion via result event.
# Returns 0 on successful result, 1 on timeout/no result.
run_claude_with_completion_detection() {
    local prompt_file="$1"
    local model="$2"
    local temp_out="$3"
    local err_log="${temp_out}.err"

    > "$temp_out"
    > "$err_log"

    # Start claude in background with stream-json output
    # Prompt piped via stdin to handle large prompts; stdout=json, stderr=separate log
    cd "$PROJECT_DIR" && cat "$prompt_file" \
        | claude -p --dangerously-skip-permissions --verbose \
            --output-format stream-json --model "$model" \
            > "$temp_out" 2>"$err_log" &
    local claude_pid=$!

    # Hard timeout watchdog (kills claude if stream-json detection fails)
    ( sleep $HARD_TIMEOUT; kill $claude_pid 2>/dev/null ) &
    local watchdog_pid=$!

    # Monitor stream-json output for the result event
    local result_received=false
    while kill -0 $claude_pid 2>/dev/null; do
        if grep -q '"type":"result"' "$temp_out" 2>/dev/null; then
            result_received=true
            # Give claude 3s to exit cleanly, then force kill
            ( sleep 3; kill $claude_pid 2>/dev/null ) &
            local killer_pid=$!
            wait $claude_pid 2>/dev/null
            kill $killer_pid 2>/dev/null
            break
        fi
        sleep 1
    done

    # Clean up watchdog
    kill $watchdog_pid 2>/dev/null
    wait $watchdog_pid 2>/dev/null
    wait $claude_pid 2>/dev/null

    # Final check: process may have exited (e.g. hook crash) after emitting the result
    # but before our polling loop caught it
    if [ "$result_received" = false ] && grep -q '"type":"result"' "$temp_out" 2>/dev/null; then
        result_received=true
    fi

    # Extract and display the result text
    local result_text
    result_text=$(grep '"type":"result"' "$temp_out" | python3 -c "
import sys, json
for line in sys.stdin:
    line = line.strip()
    if line:
        try:
            obj = json.loads(line)
            if obj.get('result'):
                print(obj['result'][:500])
                break
        except: pass
" 2>/dev/null)
    [ -n "$result_text" ] && echo "$result_text"

    if [ "$result_received" = true ]; then
        echo "  (completed via stream-json result detection)"
        rm -f "$err_log"
        return 0
    else
        # Show stderr to help diagnose failures
        if [ -s "$err_log" ]; then
            echo "  stderr output:"
            head -5 "$err_log" | sed 's/^/    /'
        fi
        echo "  (no result event received)"
        rm -f "$err_log"
        return 1
    fi
}

# Parse arguments
for arg in "$@"; do
    if [ "$arg" = "plan" ]; then
        MODE="plan"
    elif [ "$arg" = "include-tests" ]; then
        INCLUDE_TESTS=true
    elif [ "$arg" -eq "$arg" ] 2>/dev/null; then
        MAX_ITERATIONS=$arg
    fi
done

echo "=== Cascade Industries Test Suite Ralph Loop ==="
echo "Mode: $MODE"
echo "Tests: $INCLUDE_TESTS"
echo "Project: $PROJECT_DIR"
if [ $MAX_ITERATIONS -gt 0 ]; then
    echo "Max iterations: $MAX_ITERATIONS"
fi
echo ""

# Select prompt file (prompts live at project root)
if [ "$MODE" = "plan" ]; then
    PROMPT_FILE="PROMPT_plan.md"
else
    PROMPT_FILE="PROMPT_build.md"
fi

# Check prompt file exists
if [ ! -f "$PROJECT_DIR/$PROMPT_FILE" ]; then
    echo "Error: $PROMPT_FILE not found in $PROJECT_DIR"
    exit 1
fi

# Main loop
while true; do
    ITERATION=$((ITERATION + 1))
    START_EPOCH=$(date +%s)

    echo ""
    echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
    echo "  Iteration $ITERATION — $(date)"
    echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"

    # Show next bead to work on (mirrors PROMPT_build.md logic: in_progress first, then ready)
    echo ""
    IN_PROGRESS=$(cd "$PROJECT_DIR" && bd list --status=in_progress 2>/dev/null | head -1)
    if [ -n "$IN_PROGRESS" ]; then
        echo "Resuming in-progress bead:"
        echo "  $IN_PROGRESS"
    else
        echo "Next ready bead:"
        cd "$PROJECT_DIR" && bd ready 2>/dev/null | head -1 || echo "  (could not fetch beads)"
    fi
    echo ""

    # Phase 1: Build/plan with Opus 4.6
    # Uses stream-json to detect completion and kill hung process (GitHub #19060 fix)
    echo "  Phase 1: $MODE ($BUILD_MODEL)"
    set +e
    run_claude_with_completion_detection "$PROJECT_DIR/$PROMPT_FILE" "$BUILD_MODEL" "$TEMP_OUTPUT"
    BUILD_EXIT=$?
    set -e

    BUILD_ELAPSED=$(( $(date +%s) - START_EPOCH ))
    echo ""
    echo "  $MODE phase completed (exit $BUILD_EXIT, ${BUILD_ELAPSED}s)"

    # Fallback: create tracking bead if build phase crashed without creating its own beads
    if [ $BUILD_EXIT -ne 0 ]; then
        echo "  ⚠ $MODE phase exited $BUILD_EXIT — checking for untracked failures..."
        EXISTING=$(cd "$PROJECT_DIR" && bd list --status=open 2>/dev/null | grep -c "Loop iteration.*$MODE.*crash" || echo "0")
        if [ "${EXISTING}" = "0" ]; then
            cd "$PROJECT_DIR" && bd create \
                --title="Loop iteration $ITERATION $MODE phase crash (exit $BUILD_EXIT)" \
                --type=bug \
                --priority=1 \
                --labels="loop,$MODE-crash" 2>/dev/null || true
            echo "  Created fallback bead for $MODE phase failure"
        fi
    fi

    # Phase 2: Generator + grader checks (only with include-tests, only in build mode)
    # This replaces the Go-era fuzz + visual phases. It runs pytest, ruff,
    # and a determinism smoke test if the generator exists.
    if [ "$INCLUDE_TESTS" = true ] && [ "$MODE" = "build" ]; then
        echo ""
        echo "  Phase 2: Python checks (pytest, ruff, determinism smoke test)"
        CHECK_START=$(date +%s)
        set +e

        # ruff (only if config exists)
        if [ -f "$PROJECT_DIR/pyproject.toml" ] || [ -f "$PROJECT_DIR/ruff.toml" ]; then
            (cd "$PROJECT_DIR" && uv run ruff check . 2>&1 | tail -20)
            RUFF_EXIT=$?
        else
            echo "  (skipping ruff — no pyproject.toml / ruff.toml yet)"
            RUFF_EXIT=0
        fi

        # pytest (only if tests exist)
        if [ -d "$PROJECT_DIR/tests" ]; then
            (cd "$PROJECT_DIR" && uv run python -m pytest tests/ -x --tb=short 2>&1 | tail -30)
            PYTEST_EXIT=$?
        else
            echo "  (skipping pytest — no tests/ directory yet)"
            PYTEST_EXIT=0
        fi

        # Determinism smoke test: generate twice, diff outputs
        if [ -f "$PROJECT_DIR/generate_test_suite.py" ]; then
            RUN1=$(mktemp -d)
            RUN2=$(mktemp -d)
            (cd "$PROJECT_DIR" && uv run python generate_test_suite.py --output "$RUN1" 2>/dev/null)
            GEN1_EXIT=$?
            (cd "$PROJECT_DIR" && uv run python generate_test_suite.py --output "$RUN2" 2>/dev/null)
            GEN2_EXIT=$?

            if [ $GEN1_EXIT -eq 0 ] && [ $GEN2_EXIT -eq 0 ]; then
                if diff -r "$RUN1" "$RUN2" > /dev/null 2>&1; then
                    echo "  ✓ Determinism smoke test passed"
                    DETERM_EXIT=0
                else
                    echo "  ⚠ Determinism FAILED — generator emits non-identical output across runs"
                    DETERM_EXIT=1
                    EXISTING_D=$(cd "$PROJECT_DIR" && bd list --status=open 2>/dev/null | grep -c "Determinism regression" || echo "0")
                    if [ "${EXISTING_D}" = "0" ]; then
                        cd "$PROJECT_DIR" && bd create \
                            --title="Determinism regression detected in iteration $ITERATION" \
                            --type=bug \
                            --priority=0 \
                            --labels="loop,determinism" \
                            --description="generate_test_suite.py produced non-identical outputs on two consecutive runs. Violates §9 quality gate. Check for unseeded randomness, PDF timestamps (reportlab invariant flag, fpdf2 creation_date), unsorted dict iteration, or datetime.now() calls." 2>/dev/null || true
                    fi
                fi
            else
                echo "  ⚠ generate_test_suite.py failed to run (exits $GEN1_EXIT, $GEN2_EXIT)"
                DETERM_EXIT=1
            fi
            rm -rf "$RUN1" "$RUN2"
        else
            echo "  (skipping determinism smoke test — generate_test_suite.py not yet implemented)"
            DETERM_EXIT=0
        fi

        # Auto-grader self-test: gold standards must score 3/3/3/3/3 against themselves
        if [ -f "$PROJECT_DIR/scoring/auto_grader.py" ] && [ -d "$PROJECT_DIR/gold_standards" ]; then
            (cd "$PROJECT_DIR" && uv run python scoring/auto_grader.py --self-test 2>&1 | tail -20)
            GRADER_EXIT=$?
            if [ $GRADER_EXIT -ne 0 ]; then
                echo "  ⚠ auto_grader self-test failed — gold standards do not pass their own tests"
            fi
        else
            echo "  (skipping auto_grader self-test — not yet implemented)"
            GRADER_EXIT=0
        fi

        set -e
        CHECK_ELAPSED=$(( $(date +%s) - CHECK_START ))
        echo "  Checks completed in ${CHECK_ELAPSED}s (ruff=$RUFF_EXIT pytest=$PYTEST_EXIT determinism=$DETERM_EXIT grader=$GRADER_EXIT)"
    fi

    ELAPSED=$(( $(date +%s) - START_EPOCH ))
    echo ""
    echo "Iteration $ITERATION completed (total ${ELAPSED}s)"
    echo ""

    # Check for explicit exit signal (file-based)
    if [ -f "$PROJECT_DIR/.ralph-exit" ]; then
        echo "Exit signal detected (.ralph-exit file found)"
        rm -f "$PROJECT_DIR/.ralph-exit"
        break
    fi

    # Check iteration limit
    if [ $MAX_ITERATIONS -gt 0 ] && [ $ITERATION -ge $MAX_ITERATIONS ]; then
        echo "Reached maximum iterations ($MAX_ITERATIONS)"
        break
    fi

    # Small delay between iterations to avoid hammering
    sleep 2
done

echo "=== Loop completed ==="
echo "Total iterations: $ITERATION"
