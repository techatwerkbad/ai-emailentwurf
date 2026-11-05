// Výchozí nastavení doplňku pro celý tým.
// Pokud zde vyplníte webhook a token, uživatelé nemusí nic nastavovat.
// Každý uživatel si je může přepsat v panelu (Einstellungen) – uloží se do jeho schránky.
window.WERKBAD_ADDIN_CONFIG = {
  webhookUrl: "https://hook.eu2.make.com/4amgp05r4hhg8snaoot67ov1dbvx55g5",        // např. "https://hook.eu1.make.com/xxxxxxxxxxxxxxxxxxxxxxxx"
  token: "",               // sdílený tajný klíč, Make ho kontroluje v hlavičce X-Token
  timeoutMs: 170000,       // Make čeká na Webhook response max. 180 s
  maxBodyChars: 100000     // delší text e-mailu se zkrátí
};
