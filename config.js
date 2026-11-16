// Výchozí nastavení doplňku pro celý tým.
// Pokud zde vyplníte webhook a token, uživatelé nemusí nic nastavovat.
// Každý uživatel si je může přepsat v panelu (Settings) – uloží se do jeho schránky.
window.WERKBAD_ADDIN_CONFIG = {
  webhookUrl: "https://hook.eu2.make.com/b72ey5mwom2ovvs1ml2y8irtpaznguww",        // např. "https://hook.eu1.make.com/xxxxxxxxxxxxxxxxxxxxxxxx"
  badicumWebhookUrl: "",   // webhook scénáře Badicum-Email-AI-Agent (doplnit po vytvoření)
  badicumDomains: ["badicum.de"], // e-maily do těchto domén jdou agentovi Badicum
  token: "",               // sdílený tajný klíč, Make ho kontroluje v hlavičce X-Token
  timeoutMs: 170000,       // Make čeká na Webhook response max. 180 s
  maxBodyChars: 100000     // delší text e-mailu se zkrátí
};
