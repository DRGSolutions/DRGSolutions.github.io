// PowerList
//
// This file defines the list of power utilities shown in the "Power Rules" dropdown.
// It is intentionally a plain global (no bundler / no backend required).

(function () {
  const list = [
    { name: "Penn Power", rulesFile: "Penn_Power.rls" },
    { name: "Ohio Edison", rulesFile: "Ohio_Edison.rls" },
  ];

  // Expose as a stable global for app.js
  window.POWER_UTILITY_LIST = list;
})();
