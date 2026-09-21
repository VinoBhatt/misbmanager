// All writes must report failures before a dialog is closed or data is refreshed.
const nativeFetch = window.fetch.bind(window);
window.fetch = async (...args) => {
  const response = await nativeFetch(...args);
  if (!response.ok) {
    const result = await response.clone().json().catch(() => ({}));
    if (response.status === 401) location.assign('/login');
    throw new Error(result.error || `Request failed (${response.status}). Please try again.`);
  }
  return response;
};
function showFailure(message) {
  let box = document.getElementById('requestError');
  if (!box) {
    box = document.createElement('div'); box.id = 'requestError'; box.setAttribute('role','alert');
    document.body.append(box);
  }
  box.textContent = message;
  box.hidden = false;
  setTimeout(() => { box.hidden = true; }, 12000);
  document.querySelectorAll('.upload button, #saveNoteBtn').forEach(button => { button.disabled = false; });
}
window.addEventListener('unhandledrejection', event => {
  showFailure(event.reason?.message || 'Unable to complete this request. Please try again.');
  event.preventDefault();
});
