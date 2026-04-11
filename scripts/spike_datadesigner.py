"""Spike: DataDesigner with TC-04 lease seed data.

Evaluates NVIDIA NeMo DataDesigner for generating narrative-quality
lease document text from canonical seed data.

Run: uv run python scripts/spike_datadesigner.py
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

import pandas as pd

# ── 1. Build canonical seed data from the lease model ───────────────────────

def build_lease_seed_dataframe() -> pd.DataFrame:
    """Extract a DataFrame of lease facts from the canonical model."""
    # Add project root to path so we can import the generator
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

    from generator.model.build import build_model

    model = build_model()
    rows = []
    for lease in model.leases:
        rows.append({
            "lease_id": lease.lease_id,
            "entity_code": lease.entity_code,
            "lessee": lease.lessee,
            "lessor": lease.lessor,
            "description": lease.description,
            "commencement_date": str(lease.commencement_date),
            "term_months": lease.term_months,
            "monthly_base_rent": float(lease.monthly_base_rent),
            "lease_type": lease.lease_type.value,
            "short_term_exempt": lease.short_term_exempt,
            "has_amendments": len(lease.amendments) > 0,
            "amendment_count": len(lease.amendments),
        })
    return pd.DataFrame(rows)


# ── 2. Build DataDesigner config ────────────────────────────────────────────

def build_config():
    """Build a DataDesigner config that generates lease narrative text
    from canonical seed data."""
    import data_designer.config as dd

    builder = dd.DataDesignerConfigBuilder()

    # Seed the builder with the lease DataFrame
    df = build_lease_seed_dataframe()
    print(f"Seed data: {len(df)} leases, columns: {list(df.columns)}")
    print(df[["lease_id", "lessor", "description", "monthly_base_rent"]].to_string())
    print()

    builder.with_seed_dataset(dd.DataFrameSeedSource(df=df))

    # Add LLM-generated narrative column: lease description paragraph
    builder.add_column(
        dd.LLMTextColumnConfig(
            name="lease_narrative",
            model_alias="openrouter-text",
            prompt=(
                "Write a formal lease agreement opening paragraph for the following lease:\n"
                "- Lease ID: {{ lease_id }}\n"
                "- Lessee: {{ lessee }}\n"
                "- Lessor: {{ lessor }}\n"
                "- Description: {{ description }}\n"
                "- Commencement Date: {{ commencement_date }}\n"
                "- Term: {{ term_months }} months\n"
                "- Monthly Base Rent: ${{ monthly_base_rent }}\n"
                "- Lease Type: {{ lease_type }}\n\n"
                "The paragraph should read like an authentic commercial lease document. "
                "Include the exact figures provided. Do not invent additional financial terms."
            ),
        )
    )

    return builder


# ── 3. Attempt preview ──────────────────────────────────────────────────────

def run_spike():
    """Run the DataDesigner spike and document results."""
    print("=" * 72)
    print("SPIKE: DataDesigner with TC-04 Lease Seed Data")
    print("=" * 72)
    print()

    # Phase 1: Verify seed data extraction
    print("--- Phase 1: Seed Data Extraction ---")
    df = build_lease_seed_dataframe()
    print(f"Successfully extracted {len(df)} leases from canonical model")
    print()

    # Phase 2: Check provider availability
    print("--- Phase 2: Provider Availability ---")
    api_keys = {
        "NVIDIA_API_KEY": bool(os.environ.get("NVIDIA_API_KEY")),
        "OPENAI_API_KEY": bool(os.environ.get("OPENAI_API_KEY")),
        "OPENROUTER_API_KEY": bool(os.environ.get("OPENROUTER_API_KEY")),
    }
    print(f"API keys available: {api_keys}")
    any_key = any(api_keys.values())
    print()

    # Phase 3: Build config and attempt preview
    print("--- Phase 3: DataDesigner Config ---")
    try:
        builder = build_config()
        config = builder.build()
        print("Config built successfully")
        print(f"  Columns: {[c.name for c in config.columns]}")
        print()
    except Exception as e:
        print(f"Config build failed: {e}")
        print()
        return

    # Phase 4: Attempt preview (will fail without API keys)
    print("--- Phase 4: Preview Attempt ---")
    if not any_key:
        print("SKIPPED: No API keys available.")
        print("To run preview, set one of: NVIDIA_API_KEY, OPENAI_API_KEY, OPENROUTER_API_KEY")
        print()
    else:
        try:
            from data_designer.interface import DataDesigner
            designer = DataDesigner()
            preview = designer.preview(config_builder=builder, num_records=2)
            print("Preview succeeded!")
            print(preview.display_sample_record())
        except Exception as e:
            print(f"Preview failed: {e}")
        print()

    # Phase 5: Document findings
    print("=" * 72)
    print("SPIKE FINDINGS")
    print("=" * 72)
    print()
    print("1. SETUP REQUIREMENTS")
    print("   - pip install data-designer (v0.5.6)")
    print("   - Python 3.10-3.13 (compatible with our 3.11+ requirement)")
    print("   - Requires API key for at least one LLM provider:")
    print("     NVIDIA (build.nvidia.com), OpenAI, or OpenRouter")
    print("   - Default models: nvidia/nemotron-3-nano-30b-a3b (text),")
    print("     openai/gpt-oss-20b (reasoning), etc.")
    print()
    print("2. OUTPUT SHAPE")
    print("   - DataDesigner produces tabular data (pandas DataFrames)")
    print("   - Each row = one seed record + LLM-generated columns")
    print("   - For our use case: lease facts in, narrative paragraphs out")
    print("   - Supports Jinja2 template variables in prompts ({{ field }})")
    print("   - Output is per-row text, not per-document")
    print()
    print("3. DETERMINISM / CACHING IMPLICATIONS")
    print("   - DataDesigner does NOT guarantee deterministic LLM outputs")
    print("   - Temperature/top_p are configurable but LLM sampling is inherently stochastic")
    print("   - No built-in output caching mechanism")
    print("   - To satisfy our SEED=42 byte-identical requirement:")
    print("     a) Must cache generated outputs locally (artifact store)")
    print("     b) Cache keyed by: prompt hash + model + seed data hash + config")
    print("     c) Default generation path must use cache-only (no live calls)")
    print("     d) Cache warm-up is a separate explicit step")
    print()
    print("4. DEPENDENCY / PROVIDER REQUIREMENTS")
    print("   - data-designer pulls in: data-designer-config, data-designer-engine,")
    print("     prompt-toolkit, typer, pydantic, httpx, tiktoken, scipy, etc.")
    print("   - Heavy dependency tree (~50+ transitive packages)")
    print("   - Runtime requires network access to LLM provider")
    print("   - Not compatible with our 'no network calls during generation' rule")
    print()
    print("5. RECOMMENDATION: MIMIC THE PATTERN, DON'T ADOPT DIRECTLY")
    print("   Reasons to NOT adopt DataDesigner directly:")
    print("   - Adds ~50 transitive dependencies to locked stack")
    print("   - Requires LLM provider at generation time (violates offline rule)")
    print("   - No built-in caching for deterministic replay")
    print("   - Output shape (DataFrame rows) doesn't map directly to our")
    print("     document-oriented format (PDFs, XLSX, DOCX)")
    print("   - Overkill for our use case: we need ~50-100 narrative snippets,")
    print("     not scalable data pipelines")
    print()
    print("   Recommended alternative: MIMIC THE PATTERN")
    print("   - Write a thin augmentation layer that:")
    print("     a) Takes canonical facts as seed (same as DataDesigner seed)")
    print("     b) Calls any OpenAI-compatible API to generate narrative text")
    print("     c) Caches outputs keyed by (prompt_hash, model, seed_hash)")
    print("     d) Falls back to template text when cache is cold")
    print("     e) Integrates with existing formatter pipeline")
    print("   - Dependencies: just httpx or requests (already available via other deps)")
    print("   - Control: full control over caching, determinism, and output format")
    print("   - Complexity: ~200 lines vs ~50 new transitive packages")
    print()


if __name__ == "__main__":
    run_spike()
