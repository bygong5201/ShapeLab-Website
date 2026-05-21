/*
  Optional parallax motion for the blueprint background.
  Include before </body> after the .parallax-scene markup.
*/
(() => {
  const scene = document.querySelector('.parallax-scene');
  const layers = document.querySelectorAll('.parallax-layer');
  const motionQuery = window.matchMedia('(prefers-reduced-motion: reduce)');

  if (!scene || !layers.length || motionQuery.matches) return;

  let ticking = false;

  window.addEventListener('pointermove', (event) => {
    const pointerX = (event.clientX / window.innerWidth - 0.5) * 2;
    const pointerY = (event.clientY / window.innerHeight - 0.5) * 2;
    scene.style.setProperty('--mouse-x', `${(pointerX * -10).toFixed(2)}px`);
    scene.style.setProperty('--mouse-y', `${(pointerY * -8).toFixed(2)}px`);
  }, { passive: true });

  function updateParallax() {
    const scrollY = window.scrollY || window.pageYOffset;
    const maxScroll = Math.max(1, document.documentElement.scrollHeight - window.innerHeight);
    const progress = Math.min(1, Math.max(0, scrollY / maxScroll));

    layers.forEach((layer) => {
      const speed = Number(layer.dataset.parallaxSpeed || 1);
      layer.style.setProperty('--parallax-y', `${progress * speed * -90}px`);
    });

    ticking = false;
  }

  window.addEventListener('scroll', () => {
    if (!ticking) {
      requestAnimationFrame(updateParallax);
      ticking = true;
    }
  }, { passive: true });

  updateParallax();
})();
