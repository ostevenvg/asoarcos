<template>
  <div>
    <h1>Webtail + Vue</h1>
    <p>Aquí puedes mostrar los datos de Webtail</p>
  </div>
</template>

<script>
export default {
  name: "App",
  api: "http://127.0.0.1:8000/",

  data () {
    return {
      // Variable para almacenar los datos de Webtail
      data: null,
    };
  },

  mounted () {
    // Carga los datos de Webtail
    this.fetchData();
  },

  methods: {
    async fetchData () {
      const response = await fetch(this.api + "/data");
      const data = await response.json();

      this.data = data;
    },
  },
};
</script>