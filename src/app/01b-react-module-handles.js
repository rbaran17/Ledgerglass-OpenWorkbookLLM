import ne from "react";
import {
  copyModuleProperties,
  unwrapDefaultExport,
} from "./01-core-schemas-and-hashing.js";

/* React as the assistant chunk expects it. Kept apart from the core model so
   the workers can import the core model without pulling React in. */
const Gq = unwrapDefaultExport(ne),
  uEe = copyModuleProperties({ __proto__: null, default: Gq }, [ne]);

export { Gq, uEe };
