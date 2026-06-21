(function () {
  var THEME_STORAGE_KEY = 'app-theme-v2';

  function getPreferredTheme() {
    var storedTheme = localStorage.getItem(THEME_STORAGE_KEY);
    if (storedTheme === 'dark' || storedTheme === 'light') return storedTheme;
    return 'light';
  }

  function applyTheme(theme) {
    document.documentElement.setAttribute('data-theme', theme);
  }

  function syncThemeToggle(button) {
    var isDark = document.documentElement.getAttribute('data-theme') === 'dark';
    button.setAttribute('aria-label', isDark ? 'Alternar para tema claro' : 'Alternar para tema escuro');

    var icon = button.querySelector('i');
    if (!icon) return;

    if (icon.classList.contains('bi') || icon.className.indexOf('bi-') !== -1) {
      icon.classList.toggle('bi-moon-stars', !isDark);
      icon.classList.toggle('bi-sun', isDark);
      return;
    }

    icon.classList.toggle('fa-moon', !isDark);
    icon.classList.toggle('fa-sun', isDark);
  }

  applyTheme(getPreferredTheme());

  document.addEventListener('DOMContentLoaded', function () {
    var buttons = document.querySelectorAll('[data-theme-toggle], #theme-toggle');
    Array.prototype.forEach.call(buttons, function (button) {
      syncThemeToggle(button);
      button.addEventListener('click', function () {
        var current = document.documentElement.getAttribute('data-theme') === 'dark' ? 'dark' : 'light';
        var next = current === 'dark' ? 'light' : 'dark';
        applyTheme(next);
        localStorage.setItem(THEME_STORAGE_KEY, next);
        localStorage.removeItem('app-theme');
        Array.prototype.forEach.call(buttons, syncThemeToggle);
      });
    });
  });
})();
