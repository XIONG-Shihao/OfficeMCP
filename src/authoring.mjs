/** Public build-time contract surface for host guides and schema validation.
 * Runtime hosts should use the MCP CLI; the optional embedding SDK is unchanged.
 */
export {operationRegistry, VERSION} from './contracts/registry.mjs';
export {DOCX_AGENT_CONTRACT} from './formats/docx/docx-agent-contract.mjs';
export {DOCX_ACTION_REGISTRY, DOCX_ACTION_FAMILIES} from './formats/docx/docx-agent-actions.mjs';
