// Property: every generated deck emits schema-valid PresentationML.
//
// For a randomized-but-valid deck, every slide part `pptx-kit` produces
// must validate against the ECMA-376 `pml.xsd`. This is the fuzz-scale
// version of the per-feature L3 schema tests: instead of one hand-written
// slide per shape feature, it throws thousands of random shape / fill /
// text / table / bullet combinations at the validator.
//
// Like the existing L3 tests, it skips cleanly when `xmllint` is absent so
// contributors without libxml2 can still run the rest of the suite.

import { describe, expect, it } from 'vitest';
import fc from 'fast-check';
import {
  buildPresentation,
  PROPERTY_SEED,
  presentationArbitrary,
} from '../lib/arbitrary-presentation.ts';
import { getSlides, getSlideXmlString } from '../../src/api/index.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from '../lib/expect-schema-valid.ts';

// xmllint spawns one process per slide, so this property is I/O-bound;
// keep the run count lower than the pure-JS round-trip property.
const NUM_RUNS = 25;

describe('property: emitted PresentationML is schema-valid', () => {
  it.runIf(isSchemaValidationAvailable())('every slide validates against pml.xsd', () => {
    fc.assert(
      fc.property(presentationArbitrary(), (spec) => {
        const pres = buildPresentation(spec);
        const slides = getSlides(pres);
        expect(slides.length).toBe(spec.slides.length);
        for (const slide of slides) {
          expectSchemaValid(getSlideXmlString(slide), 'pml');
        }
      }),
      { seed: PROPERTY_SEED, numRuns: NUM_RUNS },
    );
  });
});
