import test from "./support/test.mjs";
import { writeDocxGuide } from "../scripts/build-agent-guide.mjs";

// agent/OfficeMCP_docx.md is generated from the contracts; a host gives it to
// its agent as is. A contract change without a regenerated guide fails here.
test("the committed agent guide matches the contracts", () => {
  writeDocxGuide({ check: true });
});
