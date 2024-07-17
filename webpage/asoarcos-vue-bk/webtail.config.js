module.exports = {
  // Ruta a la API de Webtail
  api: "https://api.webtail.io/",

  // Token de acceso a la API de Webtail
  token: "YOUR_TOKEN",

  // Proyecto de Webtail a usar
  project: "YOUR_PROJECT",

  // Entorno de Webtail a usar
  environment: "YOUR_ENVIRONMENT",

  // Elemento HTML donde se montará la aplicación Vue
  mount: "#app",

  // Componentes Vue a usar
  components: {
    // Componente principal de la aplicación
    App: require("./App.vue").default,

    // Otros componentes...
  },
};