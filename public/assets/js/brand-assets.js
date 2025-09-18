(function () {
  const HERO_LOGO_PATH = 'assets/GEP-Group_Logotipo_horizontal.png';
  const FAVICON_PATH = 'assets/GEP_Logo_Icono_GRIS.jpg';

  document.addEventListener('DOMContentLoaded', () => {
    const heroLogo = document.querySelector('.hero-logo');
    if (heroLogo) {
      heroLogo.src = HERO_LOGO_PATH;
    }

    let faviconLink = document.querySelector('link[rel="icon"]');
    if (!faviconLink) {
      faviconLink = document.createElement('link');
      faviconLink.setAttribute('rel', 'icon');
      document.head.appendChild(faviconLink);
    }

    faviconLink.setAttribute('type', 'image/jpeg');
    faviconLink.href = FAVICON_PATH;
  });
})();
