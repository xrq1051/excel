// eslint-disable-next-line no-undef, @typescript-eslint/no-unused-vars
Office.onReady((info) => {
  // eslint-disable-next-line no-undef
  document.getElementById("ok-button").onclick = () => tryCatch(sendStringToParentPage);
});

function sendStringToParentPage() {
  // eslint-disable-next-line no-undef
  const userName = document.getElementById("name-box").value;
  // eslint-disable-next-line no-undef
  Office.context.ui.messageParent(userName);
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    Console.error(error);
  }
}
