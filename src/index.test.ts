import { describe, expect, it } from 'vitest';
import { VERSION } from './index.ts';

describe('pptx-kit', () => {
  it('exports a VERSION string', () => {
    expect(typeof VERSION).toBe('string');
  });
});
