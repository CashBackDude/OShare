// Copyright (c) Microsoft Corporation. Licensed under the MIT license.

Office.onReady(() => {
  // Office.js is ready
});

function openForm(event) {
  window.open("https://forms.office.com/e/0WMwRUR02J"); // Target URL
  event.completed();
}

// Make function globally accessible
if (typeof window !== "undefined") {
  window.openForm = openForm;
}