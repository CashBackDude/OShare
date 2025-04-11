Office.initialize = function () {
  console.log("Office.js initialized");
};

function launchForm(event) {
  Office.context.ui.displayDialogAsync(
    "https://forms.office.com/e/0WMwRUR02J",
    { height: 60, width: 60 },
    function (asyncResult) {
      if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        console.error("Dialog failed:", asyncResult.error.message);
      }
    }
  );
  event.completed();
}
