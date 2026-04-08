# CRCs

## SplitPlanner

- Responsibilities:
  - Capture the current repo structure and split constraints.
  - Record the extraction and cutover workflow.
  - Define validation gates before source cleanup.
- Collaborators:
  - HistoryExtractor
  - NamingNormalizer
  - ValidationRunner
  - SourceRepoCleaner

## HistoryExtractor

- Responsibilities:
  - Extract `pitfall-clone` history into a standalone branch/repository.
  - Ensure the extracted repo root contains the former game contents directly.
- Collaborators:
  - SplitPlanner
  - ValidationRunner

## NamingNormalizer

- Responsibilities:
  - Replace legacy `pitfall-clone` and `Temple Runaway` identifiers with Dan's Dungeon naming where user-facing.
  - Keep gameplay code and runtime behavior otherwise unchanged.
- Collaborators:
  - HistoryExtractor
  - ValidationRunner

## ValidationRunner

- Responsibilities:
  - Verify install, build, test, and browser shell identity in the new repo.
  - Confirm source repo references no longer depend on the game after cleanup.
- Collaborators:
  - NamingNormalizer
  - SourceRepoCleaner

## SourceRepoCleaner

- Responsibilities:
  - Remove `pitfall-clone` from `danmade-playground` after validation passes.
  - Preserve split documentation in the source repo.
- Collaborators:
  - ValidationRunner
  - SplitPlanner
