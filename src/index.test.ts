import { describe, expect, it } from 'vitest';
import { VERSION } from './index.ts';

describe('@office-kit/pptx', () => {
  it('exports a VERSION string', () => {
    expect(typeof VERSION).toBe('string');
  });
});
