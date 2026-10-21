// `@office-kit/pptx-editor/headless`: a small set of edits a language model
// applies to a presentation without mounting the editor, for example on a
// server. It has no DOM dependency and runs in Node as well as the browser.

export {
  AgentEditError,
  agentEditSchema,
  applyAgentEdits,
  decodeAgentEdits,
  readForAgent,
  type DecodedAgentEdits,
} from './core/agent-edits.ts';
export type { JsonSchema, JsonValue } from './core/agent-tools.ts';
