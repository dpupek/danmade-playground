# Edge Cases

## History Preservation

- Risk: copying files directly would lose commit history.
- Mitigation: extract with `git subtree split --prefix=pitfall-clone` so the new repo history is rooted in the game path.

## Naming Drift

- Risk: the extracted repo could still mix `pitfall-clone`, `Temple Runaway`, and `Dan's Dungeon`.
- Mitigation: update package metadata, README, HTML title, docs, and generation prompts/scripts in one pass.

## Path References

- Risk: docs and scripts may contain hard-coded source repo paths.
- Mitigation: rewrite setup paths and output references to the new `E:\Sandbox\dungeon-dan` root.

## Source Repo Cleanup

- Risk: removing the folder before validation could leave the game in a broken state or force a second extraction.
- Mitigation: validate the new repo first, then remove `pitfall-clone` from the source repo.

## Generated and Ignored Files

- Risk: build outputs, `node_modules`, or generated art could pollute the new repo state.
- Mitigation: keep existing ignore rules and validate from a fresh install/build workflow.
