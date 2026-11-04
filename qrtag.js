/* Private QR format. A label's QR holds a scrambled value (AQ1...) that only this app can decode and check.
   A customer scanning it with an ordinary phone camera sees meaningless text.
   Change SECRET once, BEFORE printing labels, and never change it afterwards (printed labels would stop working).
   This keeps casual scanners out; it is not bank-grade security. */
(function(){
  const SECRET = "atig-private-qr-2026";
  const h53 = (str, seed = 0) => {
    let h1 = 0xdeadbeef ^ seed, h2 = 0x41c6ce57 ^ seed;
    for (let i = 0, ch; i < str.length; i++){ ch = str.charCodeAt(i); h1 = Math.imul(h1 ^ ch, 2654435761); h2 = Math.imul(h2 ^ ch, 1597334677); }
    h1 = Math.imul(h1 ^ (h1 >>> 16), 2246822507) ^ Math.imul(h2 ^ (h2 >>> 13), 3266489909);
    h2 = Math.imul(h2 ^ (h2 >>> 16), 2246822507) ^ Math.imul(h1 ^ (h1 >>> 13), 3266489909);
    return 4294967296 * (2097151 & h2) + (h1 >>> 0);
  };
  const tagOf = s => (h53(SECRET + "|" + s) % 1679616).toString(36).padStart(4, "0");
  window.GQS_QR = {
    encode(serial){
      const b = new TextEncoder().encode(String(serial)); let hex = "";
      for (let i = 0; i < b.length; i++) hex += ((b[i] ^ (h53(SECRET + i) & 255)) + 256).toString(16).slice(1);
      return "AQ1" + hex + tagOf(serial);
    },
    decode(text){
      const t = String(text).trim();
      if (!/^AQ1[0-9a-f]{2,}[0-9a-z]{4}$/.test(t)) return null;
      const body = t.slice(3, -4); if (body.length % 2) return null;
      const b = new Uint8Array(body.length / 2);
      for (let i = 0; i < b.length; i++) b[i] = parseInt(body.substr(i * 2, 2), 16) ^ (h53(SECRET + i) & 255);
      let s; try{ s = new TextDecoder("utf-8", { fatal:true }).decode(b); }catch(e){ return null; }
      return tagOf(s) === t.slice(-4) ? s : null;
    }
  };
})();
