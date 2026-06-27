type ShortcutKeyEvent = Pick<
  KeyboardEvent,
  "altKey" | "ctrlKey" | "key" | "metaKey" | "shiftKey" | "target"
>;

export function isEditableShortcutTarget(target: EventTarget | null): boolean {
  if (typeof HTMLElement === "undefined" || !(target instanceof HTMLElement)) {
    return false;
  }

  if (target.isContentEditable) {
    return true;
  }

  return (
    target instanceof HTMLInputElement ||
    target instanceof HTMLSelectElement ||
    target instanceof HTMLTextAreaElement
  );
}

export function shouldCloseWorkbookSearchOnEscape(
  event: ShortcutKeyEvent,
  options: {
    hasSearchQuery: boolean;
    isFormulaInputFocused: boolean;
    isSearchOpen: boolean;
  },
): boolean {
  if (
    !options.isSearchOpen ||
    !options.hasSearchQuery ||
    event.key !== "Escape" ||
    event.altKey ||
    event.ctrlKey ||
    event.metaKey ||
    event.shiftKey
  ) {
    return false;
  }

  return !options.isFormulaInputFocused && !isEditableShortcutTarget(event.target);
}
