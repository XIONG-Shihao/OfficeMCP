import test from "./support/test.mjs";
import assert from "node:assert/strict";
import {
  DOCX_AGENT_ACTIONS,
} from "../src/formats/docx/docx-agent-actions.mjs";
import {
  LIMITED_ACTIONS,
  QUALIFICATION_GATES,
  QUALIFIED_ACTIONS,
  evidenceFor,
} from "./docx-qualification-data.mjs";

test("DOCX action qualification groups partition the registry truthfully", () => {
  const classified = new Set([...QUALIFIED_ACTIONS, ...LIMITED_ACTIONS]);
  const experimental = DOCX_AGENT_ACTIONS.filter((name) => !classified.has(name));
  assert.equal(QUALIFIED_ACTIONS.length, 50);
  assert.equal(experimental.length, 27);
  assert.equal(LIMITED_ACTIONS.length, 2);
  const grouped = [...QUALIFIED_ACTIONS, ...experimental, ...LIMITED_ACTIONS];
  assert.equal(new Set(grouped).size, grouped.length);
  assert.deepEqual([...grouped].sort(), [...DOCX_AGENT_ACTIONS].sort());
  for (const name of QUALIFIED_ACTIONS)
    assert.deepEqual(
      Object.keys(evidenceFor(name)).sort(),
      [...QUALIFICATION_GATES].sort(),
      `${name} cannot be fully qualified without all seven evidence gates`,
    );
});
