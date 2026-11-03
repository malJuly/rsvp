/*
 * The opening is real-looking footage and plain type, nothing more: an unlit lamp in a dark room.
 * It lights by itself in his memory, the verse appears, then dawn over still water, and his name.
 * The clips were generated in Google Flow; see public/assets. Timings follow the clips themselves.
 */
import Music from './music.js';

const Intro = (function () {
  const LIGHT_AFTER_MS = 1500;
  const REFUSED_HOLD_MS = 4000;

/** Runs the opening and resolves once the page is showing. */
  function play(root) {
    return new Promise((resolve) => {
      root.innerHTML = `
        <div class="film">
          <video class="film-clip lake" src="assets/lake.mp4" poster="assets/lake-poster.jpg" muted playsinline preload="auto"></video>
          <video class="film-clip lamp" src="assets/lamp.mp4" poster="assets/lamp-poster.jpg" muted playsinline preload="auto"></video>
          <p class="film-text film-om show">Om Shanti</p>
          <div class="film-scrim"></div>
          <div class="film-text film-verse">
            <p class="verse-a">You may have left this world,</p>
            <p class="verse-b">but you will never leave our hearts.</p>
          </div>
          <div class="film-text film-caption">
            <p class="remember">A gathering in loving remembrance of</p>
            <p class="name">Dr Jayaraman a/l Munusamy</p>
            <p class="span">22nd March 1938 – 30th October 2026</p>
          </div>
          <button type="button" class="gate-skip" aria-label="Skip the opening">Skip</button>
        </div>`;

      const q = (s) => root.querySelector(s);
      const film = q('.film');
      const lamp = q('.lamp');
      const lake = q('.lake');
      let done = false;
      let leaving = false;
      let lakeStarted = false;
      let autoTimer = 0;
      let safetyTimer = 0;

      function finish() {
        if (done) return;
        done = true;
        clearTimeout(autoTimer);
        clearTimeout(safetyTimer);
        lamp.pause();
        lake.pause();
        root.innerHTML = '';
        document.documentElement.classList.remove('play-intro');
        resolve();
      }

      function leave(ms) {
        if (leaving) return;
        leaving = true;
        const out = film.animate([{ opacity: 1 }, { opacity: 0 }], { duration: ms, easing: 'ease-in-out', fill: 'forwards' });
        out.finished.then(finish).catch(finish);
      }

      // A clip can report its time once more after the film has been removed.
      const show = (selector) => q(selector)?.classList.add('show');
      const hide = (selector) => q(selector)?.classList.remove('show');

      // Timed from the clips, so a slow connection delays the words along with the picture.
      lamp.addEventListener('timeupdate', () => {
        const t = lamp.currentTime;
        if (t >= 2.4) {
          film.classList.add('reading');
          show('.verse-a');
        }
        if (t >= 4) show('.verse-b');
        if (t >= 6.8 && !lakeStarted) {
          lakeStarted = true;
          lake.play().catch(() => leave(800));
          film.classList.add('dawn');
        }
      });
      lake.addEventListener('timeupdate', () => {
        const t = lake.currentTime;
        if (t >= 1.8 && t < 5) show('.film-caption');
        // His name is gone before the dissolve, so it never lies across the page's own.
        if (t >= 5) hide('.film-caption');
        if (t >= 6.4) leave(1500);
      });
      lake.addEventListener('ended', () => leave(800));
      [lamp, lake].forEach((clip) => clip.addEventListener('error', () => leave(600)));

      function light() {
        hide('.film-om');
        lamp.play().catch(() => {
          // Some phones refuse to play by themselves, in Low Power Mode for one. The still lamp and the verse stand in.
          film.classList.add('reading');
          show('.verse-a');
          show('.verse-b');
          setTimeout(() => leave(1500), REFUSED_HOLD_MS);
        });
        // If the footage can't play at all, a guest should never be left waiting in the dark.
        safetyTimer = setTimeout(() => leave(800), 30000);
      }

      Music.playWhenAllowed(4000);
      autoTimer = setTimeout(light, LIGHT_AFTER_MS);

      q('.gate-skip').addEventListener('click', () => {
        clearTimeout(autoTimer);
        clearTimeout(safetyTimer);
        leave(450);
      });
    });
  }

  return { play };
})();

export default Intro;
