import type { ChartAxisScaling, ChartSecondaryValueAxis } from './types.ts';

/** The `ChartAxisScaling` fields written flat on a secondary value axis. */
export const flatSecondaryScaling = (axis: ChartSecondaryValueAxis): ChartAxisScaling => ({
  ...(axis.min !== undefined ? { min: axis.min } : {}),
  ...(axis.max !== undefined ? { max: axis.max } : {}),
  ...(axis.majorUnit !== undefined ? { majorUnit: axis.majorUnit } : {}),
  ...(axis.minorUnit !== undefined ? { minorUnit: axis.minorUnit } : {}),
  ...(axis.numberFormat !== undefined ? { numberFormat: axis.numberFormat } : {}),
  ...(axis.logBase !== undefined ? { logBase: axis.logBase } : {}),
  ...(axis.displayUnits !== undefined ? { displayUnits: axis.displayUnits } : {}),
  ...(axis.displayUnitsLabel !== undefined ? { displayUnitsLabel: axis.displayUnitsLabel } : {}),
});

/**
 * The secondary axis's scaling from whichever form it was written in: the
 * flat fields, or the deprecated `scaling` object. Validation rejects a spec
 * that sets both, so at most one of them is non-empty here.
 */
export const secondaryAxisScaling = (
  axis: ChartSecondaryValueAxis | undefined,
): ChartAxisScaling | undefined => {
  if (axis === undefined) return undefined;
  if (axis.scaling !== undefined) return axis.scaling;
  const flat = flatSecondaryScaling(axis);
  return Object.keys(flat).length > 0 ? flat : undefined;
};
