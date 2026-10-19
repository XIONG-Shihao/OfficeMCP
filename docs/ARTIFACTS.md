# Portable saved artifact descriptors

The ordinary OfficeMCP CLI saves into its configured output root. Save and export return the output path, SHA-256 revision and `data.artifactDescriptor`, pointing to an adjacent `<output>.artifact.json` file. This behavior is the same for every supported client; it does not depend on an app, harness or model.

The descriptor is an optional, versioned product convention, **not an MCP protocol standard or negotiated extension**. A client can use the saved file directly. A host that wants source revision checks can consume the descriptor without importing OfficeMCP code.

```json
{
  "schema": "mcp-artifact/1",
  "sha256": "<64 lowercase hexadecimal characters>",
  "inputs": [
    {
      "uri": "file:///absolute/canonical/input.docx",
      "sha256": "<original source SHA-256>"
    }
  ]
}
```

`sha256` identifies the exact saved output bytes. Each input identifies the original source snapshot for the edited document, rather than a private intermediate candidate. Generated documents have an empty input array. Text supplied by a caller is not automatically traced to another file; callers or hosts must record additional workflow dependencies explicitly.

The descriptor contains no app entry IDs, actor identities, run IDs, authentication tokens or host callbacks. For example, a host can stage revision A of `contract.docx`, let OfficeMCP edit that snapshot, and later reject the output if the live source has changed to revision B. A retained OfficeMCP handle continues to name revision A even when a new host run stages revision B.

Descriptors are bounded to 64 KiB and published without overwriting existing metadata. Recovery after a process crash between document publication and descriptor publication recreates the descriptor from the durable operation record before reporting a completed receipt. A conflicting descriptor refuses completion; callers must inspect an uncertain operation rather than assume the saved output was rolled back. Old durable operations predating this convention may replay without a descriptor.

A descriptor is producer-supplied provenance, not a cryptographic proof of causal dependency or a sandbox. Hosts must restrict producer filesystem access, validate output hashes and admitted snapshot identities, check live source revisions and dirty owners, and commit through their own storage service. A malicious producer could misrepresent its inputs; applications needing stronger trust must control the producer or record dependencies independently. MCP does not define the host's workspace commit or dirty-editor policy.
