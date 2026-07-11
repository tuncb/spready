interface DialogClickEvent {
  currentTarget: EventTarget | null;
  target: EventTarget | null;
}

export function isDialogBackdropClick(event: DialogClickEvent): boolean {
  return event.target !== null && event.target === event.currentTarget;
}
