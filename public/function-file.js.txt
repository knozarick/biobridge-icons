/* global Office */
Office.onReady(() => {});
function onOpenTaskPane(event) {
  Office.addin.showAsTaskpane();
  event.completed();
}
window.onOpenTaskPane = onOpenTaskPane;
