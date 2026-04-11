"""Augmentation feature gate and cached artifact store.

Augmentation is OFF by default.  When disabled, no provider calls are made,
no cache is read or written, and generator output is byte-identical to the
pre-augmentation baseline.

When enabled, cached LLM outputs can enrich generated narrative data.
Cache misses either trigger a live LLM call (warm_on_miss=True) or fall
back to baseline template text (warm_on_miss=False).
"""

from __future__ import annotations

import hashlib
import json
import logging
from dataclasses import dataclass
from pathlib import Path
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from generator.config import AugmentationConfig

__version__ = "1.0"

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Data transfer objects
# ---------------------------------------------------------------------------

@dataclass
class AugmentationRequest:
    """Describes one augmentation prompt to be resolved against the cache."""

    prompt_template: str
    template_vars: dict
    model_alias: str
    seed_namespace: str


@dataclass
class WarmResult:
    """Summary of a cache warm-up run."""

    hits: int
    misses: int
    total_tokens: int
    estimated_cost_usd: float


# ---------------------------------------------------------------------------
# Cache key computation
# ---------------------------------------------------------------------------

def _compute_cache_key(
    scenario_hash: str,
    model_alias: str,
    prompt_text: str,
    seed_namespace: str,
) -> str:
    """Compute the SHA-256 cache key from the 5-component composite.

    The tool_version is always included as the module-level ``__version__``.
    """
    hasher = hashlib.sha256()
    hasher.update(scenario_hash.encode())
    hasher.update(__version__.encode())
    hasher.update(model_alias.encode())
    hasher.update(hashlib.sha256(prompt_text.encode()).hexdigest().encode())
    hasher.update(seed_namespace.encode())
    return hasher.hexdigest()


# ---------------------------------------------------------------------------
# AugmentationCache
# ---------------------------------------------------------------------------

class AugmentationCache:
    """Versioned cache for LLM-generated narrative artifacts.

    When augmentation is disabled (the default), all public methods are
    safe no-ops that return ``None`` or empty results.  This lets callers
    unconditionally call ``cache.get(...)`` without checking the feature
    flag at every call site.
    """

    def __init__(
        self,
        aug_config: AugmentationConfig,
        scenario_hash: str = "",
    ) -> None:
        self._config = aug_config
        self._scenario_hash = scenario_hash
        self._cache_dir = Path(aug_config.cache_dir) if aug_config.enabled else None
        self._manifest: dict | None = None

        if aug_config.enabled and not aug_config.model:
            raise ValueError(
                "augmentation.model is required when augmentation is enabled"
            )

        if self._cache_dir is not None and self._cache_dir.exists():
            self._load_manifest()

    @property
    def enabled(self) -> bool:
        return self._config.enabled

    # -- public API -----------------------------------------------------------

    def get(
        self,
        prompt: str,
        model_alias: str,
        seed_namespace: str,
    ) -> str | None:
        """Look up a cached augmentation output.

        Returns ``None`` when augmentation is disabled or on cache miss.
        """
        if not self._config.enabled:
            return None

        key = _compute_cache_key(
            self._scenario_hash, model_alias, prompt, seed_namespace,
        )
        artifact_dir = self._cache_dir / "artifacts" / key  # type: ignore[union-attr]
        output_path = artifact_dir / "output.json"

        if not output_path.exists():
            logger.debug("augmentation cache miss: %s/%s", seed_namespace, key[:12])
            return None

        with open(output_path) as f:
            data = json.load(f)
        return data.get("text")

    def put(
        self,
        prompt: str,
        model_alias: str,
        seed_namespace: str,
        output: str,
        token_usage: dict | None = None,
    ) -> str:
        """Store an augmentation output in the cache.  Returns the artifact hash."""
        if not self._config.enabled:
            return ""

        key = _compute_cache_key(
            self._scenario_hash, model_alias, prompt, seed_namespace,
        )
        artifact_dir = self._cache_dir / "artifacts" / key  # type: ignore[union-attr]
        artifact_dir.mkdir(parents=True, exist_ok=True)

        # Write output
        output_path = artifact_dir / "output.json"
        with open(output_path, "w") as f:
            json.dump({"text": output}, f, indent=2, sort_keys=True)
            f.write("\n")

        # Write metadata
        meta_path = artifact_dir / "meta.json"
        with open(meta_path, "w") as f:
            json.dump(
                {
                    "scenario_hash": self._scenario_hash,
                    "model_alias": model_alias,
                    "seed_namespace": seed_namespace,
                    "tool_version": __version__,
                    "token_usage": token_usage or {},
                },
                f,
                indent=2,
                sort_keys=True,
            )
            f.write("\n")

        self._update_manifest(key, model_alias, seed_namespace, token_usage)
        return key

    def warm(self, prompts: list[AugmentationRequest]) -> WarmResult:
        """Warm cache for all given prompts.  Calls LLM for misses only.

        Not yet implemented — returns a result with zero misses when
        augmentation is disabled.
        """
        if not self._config.enabled:
            return WarmResult(hits=0, misses=0, total_tokens=0, estimated_cost_usd=0.0)

        hits = 0
        misses = 0
        for req in prompts:
            cached = self.get(req.prompt_template, req.model_alias, req.seed_namespace)
            if cached is not None:
                hits += 1
            else:
                misses += 1
                # Live LLM call would go here when provider integration is added
                logger.info(
                    "augmentation cache cold miss (no provider): %s",
                    req.seed_namespace,
                )

        return WarmResult(
            hits=hits,
            misses=misses,
            total_tokens=0,
            estimated_cost_usd=0.0,
        )

    def prune(self) -> int:
        """Remove orphaned artifacts.  Returns count of removed entries."""
        if not self._config.enabled or self._cache_dir is None:
            return 0
        # Placeholder — full implementation when provider integration lands
        return 0

    @property
    def manifest(self) -> dict:
        if self._manifest is None:
            return {"version": 1, "tool_version": __version__, "entries": {}}
        return self._manifest

    # -- internal helpers -----------------------------------------------------

    def _load_manifest(self) -> None:
        assert self._cache_dir is not None
        manifest_path = self._cache_dir / "manifest.json"
        if manifest_path.exists():
            with open(manifest_path) as f:
                self._manifest = json.load(f)

    def _update_manifest(
        self,
        key: str,
        model_alias: str,
        seed_namespace: str,
        token_usage: dict | None,
    ) -> None:
        assert self._cache_dir is not None
        if self._manifest is None:
            self._manifest = {
                "version": 1,
                "tool_version": __version__,
                "entries": {},
            }
        self._manifest["entries"][key] = {
            "scenario_hash": self._scenario_hash,
            "model_alias": model_alias,
            "seed_namespace": seed_namespace,
            "artifact_path": f"artifacts/{key}/output.json",
            "token_usage": token_usage or {},
        }
        manifest_path = self._cache_dir / "manifest.json"
        with open(manifest_path, "w") as f:
            json.dump(self._manifest, f, indent=2, sort_keys=True)
            f.write("\n")
