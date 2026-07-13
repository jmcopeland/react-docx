import { describe, expect, it } from "vitest";
import {
  EDITOR_HISTORY_LIMIT,
  createTransactionalEditorStateReducer,
  type TransactionalEditorState,
} from "../../packages/react-viewer/src/editor-transaction-state";

interface TestModel {
  id: number;
  maxOffset: number;
}

interface TestSelection {
  offset: number;
}

interface TestRange {
  start: number;
  end: number;
}

interface TestPendingStyle {
  bold?: boolean;
}

type TestState = TransactionalEditorState<
  TestModel,
  TestSelection,
  TestRange,
  TestPendingStyle
>;

const clamp = (value: number, max: number): number =>
  Math.max(0, Math.min(max, value));

const reduce = createTransactionalEditorStateReducer<
  TestModel,
  TestSelection,
  TestRange,
  TestPendingStyle
>({
  cloneSelection: (selection) => ({ ...selection }),
  cloneTextRange: (range) => (range ? { ...range } : undefined),
  clonePendingStyle: (style) => (style ? { ...style } : undefined),
  sameSelection: (a, b) => a.offset === b.offset,
  sameTextRange: (a, b) =>
    a === b || (a?.start === b?.start && a?.end === b?.end),
  samePendingStyle: (a, b) => a?.bold === b?.bold,
  normalizeCursor: (model, selection, activeTextRange) => ({
    selection: {
      offset: clamp(selection.offset, model.maxOffset),
    },
    activeTextRange: activeTextRange
      ? {
          start: clamp(
            Math.min(activeTextRange.start, activeTextRange.end),
            model.maxOffset
          ),
          end: clamp(
            Math.max(activeTextRange.start, activeTextRange.end),
            model.maxOffset
          ),
        }
      : undefined,
  }),
});

function state(id = 0, maxOffset = 10): TestState {
  return {
    model: { id, maxOffset },
    selection: { offset: 1 },
    activeTextRange: { start: 1, end: 1 },
    pendingRunStyle: undefined,
    history: {
      past: [],
      future: [],
    },
  };
}

describe("transactional editor state", () => {
  it("commits the model, normalized cursor, pending style, and history together", () => {
    const previous = state();
    const next = reduce(previous, {
      type: "commit",
      patch: {
        model: { id: 1, maxOffset: 4 },
        selection: { offset: 99 },
        activeTextRange: { start: 8, end: 3 },
        pendingRunStyle: { bold: true },
      },
    });

    expect(next).toEqual({
      model: { id: 1, maxOffset: 4 },
      selection: { offset: 4 },
      activeTextRange: { start: 3, end: 4 },
      pendingRunStyle: { bold: true },
      history: {
        past: [
          {
            model: previous.model,
            selection: { offset: 1 },
            activeTextRange: { start: 1, end: 1 },
          },
        ],
        future: [],
      },
    });
    expect(previous.model.id).toBe(0);
    expect(previous.history.past).toHaveLength(0);
  });

  it("suppresses history pushes while invalidating a stale redo branch", () => {
    const previous = state();
    const futureSnapshot = {
      model: { id: 9, maxOffset: 10 },
      selection: { offset: 2 },
      activeTextRange: { start: 2, end: 2 },
    };
    previous.history.future = [futureSnapshot];

    const next = reduce(previous, {
      type: "commit",
      patch: { model: { id: 1, maxOffset: 10 } },
      pushHistory: false,
    });

    expect(next.model.id).toBe(1);
    expect(next.history.past).toBe(previous.history.past);
    expect(next.history.future).toEqual([]);
  });

  it("keeps the redo branch for cursor-only commits", () => {
    const previous = state();
    const futureSnapshot = {
      model: { id: 9, maxOffset: 10 },
      selection: { offset: 2 },
      activeTextRange: { start: 2, end: 2 },
    };
    previous.history.future = [futureSnapshot];

    const next = reduce(previous, {
      type: "commit",
      patch: { selection: { offset: 4 } },
      pushHistory: false,
    });

    expect(next.selection.offset).toBe(4);
    expect(next.history).toBe(previous.history);
    expect(next.history.future).toEqual([futureSnapshot]);
  });

  it("restores normalized cursors on undo and redo and clears pending style", () => {
    const initial = state();
    const edited = reduce(initial, {
      type: "commit",
      patch: {
        model: { id: 1, maxOffset: 3 },
        selection: { offset: 3 },
        activeTextRange: { start: 2, end: 3 },
        pendingRunStyle: { bold: true },
      },
    });

    const undone = reduce(edited, { type: "undo" });
    expect(undone.model.id).toBe(0);
    expect(undone.selection).toEqual({ offset: 1 });
    expect(undone.activeTextRange).toEqual({ start: 1, end: 1 });
    expect(undone.pendingRunStyle).toBeUndefined();
    expect(undone.history.future[0]).toMatchObject({
      model: edited.model,
      selection: { offset: 3 },
      activeTextRange: { start: 2, end: 3 },
    });

    const redone = reduce(undone, { type: "redo" });
    expect(redone.model.id).toBe(1);
    expect(redone.selection).toEqual({ offset: 3 });
    expect(redone.activeTextRange).toEqual({ start: 2, end: 3 });
    expect(redone.pendingRunStyle).toBeUndefined();
  });

  it("clears the redo branch after a new recorded model commit", () => {
    const edited = reduce(state(), {
      type: "commit",
      patch: { model: { id: 1, maxOffset: 10 } },
    });
    const undone = reduce(edited, { type: "undo" });
    expect(undone.history.future).toHaveLength(1);

    const branched = reduce(undone, {
      type: "commit",
      patch: { model: { id: 2, maxOffset: 10 } },
    });
    expect(branched.history.future).toHaveLength(0);
    expect(branched.history.past).toHaveLength(1);
    expect(branched.history.past[0]?.model.id).toBe(0);
  });

  it("bounds recorded history to the latest 100 snapshots", () => {
    let current = state();
    for (let id = 1; id <= EDITOR_HISTORY_LIMIT + 5; id += 1) {
      current = reduce(current, {
        type: "commit",
        patch: { model: { id, maxOffset: 10 } },
      });
    }

    expect(current.history.past).toHaveLength(EDITOR_HISTORY_LIMIT);
    expect(current.history.past[0]?.model.id).toBe(5);
    expect(current.history.past.at(-1)?.model.id).toBe(104);
  });
});
