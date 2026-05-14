# Fixture provenance

Every fixture checked into `test/fixtures/` MUST appear in this table. Origin
and scrubbing matter for OSS license cleanliness and are impossible to
reconstruct after the fact.

When you add a fixture:

1. Add one row below. Keep paths relative to `test/fixtures/`.
2. State the **source**: who created the file, with what tool. "Hand-authored"
   is acceptable for minimal fixtures.
3. State the **license**: MIT / CC0 / your own / etc. If the file is a
   template downloaded from somewhere, that source's license applies.
4. State any **scrubbing** done: PII removed, brand-specific content swapped,
   embedded media replaced with public-domain stand-ins. "None" if untouched.
5. State the **feature(s) targeted** in one short phrase — what is this
   fixture supposed to exercise?

| Path | Source | License | Scrubbing | Targets |
|---|---|---|---|---|
| _(none yet — first fixtures land in P1)_ | | | | |
