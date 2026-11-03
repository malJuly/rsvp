import Art from '../lib/art.js';
import LotusRule from './LotusRule.jsx';

const flowers = { __html: Art.portraitFlowers(200) };

export default function Hero() {
  return (
    <header className="hero">
      <p className="om">Om Shanti</p>
      <div className="portrait" id="portrait">
        <div className="portrait-oval">
          <img
            src="assets/portrait.webp"
            alt="Portrait of Dr Jayaraman a/l Munusamy, smiling, in a dark suit and glasses"
            width="432"
            height="1280"
          />
        </div>
        <span style={{ display: 'contents' }} dangerouslySetInnerHTML={flowers} />
      </div>
      <p className="script gathering">A Gathering in Loving Remembrance</p>
      <p className="honouring amp">&amp;</p>
      <p className="honouring">Honouring the Life and Legacy of</p>
      <h1>Dr Jayaraman a/l Munusamy</h1>
      <LotusRule />
      <p className="dates">22nd March 1938 – 30th October 2026</p>
    </header>
  );
}
