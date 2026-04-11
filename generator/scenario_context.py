"""ScenarioContext — deterministic seed namespaces for the test suite generator.

Provides isolated, reproducible PRNG streams keyed by name.  Adding a new
namespace never perturbs existing streams because child seeds are derived via
stable hashing (HMAC-SHA256), not Python's built-in hash().
"""

from __future__ import annotations

import hashlib
import hmac
import struct
from typing import TYPE_CHECKING

import numpy as np
from numpy.random import Generator

if TYPE_CHECKING:
    pass


class ScenarioContext:
    """Root context holding the base seed and producing named RNG streams.

    Usage::

        ctx = ScenarioContext(seed=42)
        rng_gl = ctx.named_rng("general_ledger")
        rng_ar = ctx.named_rng("accounts_receivable")

        # Child contexts for deeper isolation
        child = ctx.child("tc_08")
        rng_08 = child.named_rng("lease_data")

    Guarantees:
    - Same (base_seed, name) always produces the same stream.
    - Adding new names does not affect existing streams.
    - child_seed() produces a new ScenarioContext with a derived base seed.
    """

    def __init__(self, seed: int) -> None:
        self._base_seed = seed

    @property
    def base_seed(self) -> int:
        """The base seed this context was created with."""
        return self._base_seed

    def _derive_seed(self, name: str) -> int:
        """Derive a 64-bit seed from (base_seed, name) using HMAC-SHA256.

        The key is the base seed encoded as 8 bytes (big-endian int64).
        The message is the namespace name encoded as UTF-8.
        We take the first 8 bytes of the HMAC digest and interpret them as
        an unsigned 64-bit integer, then mask to 63 bits to stay within
        numpy's SeedSequence range (non-negative).
        """
        key = struct.pack(">q", self._base_seed)
        digest = hmac.new(key, name.encode("utf-8"), hashlib.sha256).digest()
        raw = struct.unpack(">Q", digest[:8])[0]
        # Mask to 63 bits so the seed is always non-negative
        return raw & 0x7FFFFFFFFFFFFFFF

    def named_rng(self, name: str) -> Generator:
        """Return a numpy Generator seeded deterministically by *name*.

        Each unique name produces an independent stream.  Calling this
        method multiple times with the same name returns a *new* Generator
        at the start of the same stream (not a continued one).
        """
        seed = self._derive_seed(name)
        return np.random.default_rng(np.random.SeedSequence(seed))

    def child_seed(self, name: str) -> int:
        """Derive a stable child seed integer for *name*.

        Useful when you need a raw int seed (e.g. for Faker or stdlib random)
        rather than a numpy Generator.
        """
        return self._derive_seed(name)

    def child(self, name: str) -> "ScenarioContext":
        """Create a child ScenarioContext with a derived base seed.

        The child context is fully isolated — its named_rng() streams are
        independent of the parent's streams.
        """
        return ScenarioContext(seed=self._derive_seed(name))
