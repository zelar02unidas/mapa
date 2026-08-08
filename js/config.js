/* =====================================================================
   CONFIGURACAO DO PROJETO
   ---------------------------------------------------------------------
   Preencha os dados abaixo com o seu repositorio no GitHub para que o
   site carregue a planilha "ao vivo" (a cada atualizacao da planilha
   no GitHub, o site mostra os dados novos automaticamente).

   1) Crie um repositorio publico no GitHub (ex.: "mapa-zelar").
   2) Suba a planilha "Mapa - Zelar.xlsx" NA RAIZ do repositorio.
   3) Habilite o GitHub Pages (Settings > Pages > Branch main).
   4) Edite as duas variaveis abaixo.
   ===================================================================== */

var CONFIG = {
  // Usuario ou organizacao dona do repositorio (ex.: "joao-silva")
  GITHUB_USER: "zelar02unidas",

  // Nome do repositorio (ex.: "mapa-zelar")
  GITHUB_REPO: "mapa",

  // Branch onde a planilha esta (normalmente "main")
  GITHUB_BRANCH: "main",

  // Nome exato do arquivo da planilha na raiz do repositorio
  XLSX_NAME: "Mapa - Zelar.xlsx"
};

/* Nao edite abaixo. As URLs tentadas na ordem (raw GitHub e jsDelivr). */
var CONFIG_URLS = (function () {
  var enc = encodeURIComponent(CONFIG.XLSX_NAME);
  var base = CONFIG.GITHUB_USER + "/" + CONFIG.GITHUB_REPO;
  return [
    "https://raw.githubusercontent.com/" + base + "/" + CONFIG.GITHUB_BRANCH + "/" + enc,
    "https://cdn.jsdelivr.net/gh/" + base + "@" + CONFIG.GITHUB_BRANCH + "/" + enc
  ];
})();
