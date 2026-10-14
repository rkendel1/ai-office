import type { RequirementChecker } from "@loai/core";

export function supports(set: string, version: string): boolean {
  try {
    return !!Office?.context?.requirements?.isSetSupported(set, version);
  } catch {
    return false;
  }
}

export const officeRequirements: RequirementChecker = { isSetSupported: supports };
