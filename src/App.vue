@@ -1,61 +0,0 @@
<template>
  <div id="app">
    <div id="content">
      <div id="content-header">
          <div class="padding">
          <h1>Welcome</h1>
          </div>
      </div>
      <div id="content-main">
          <div class="padding">
            <p>Choose the button below to set the color of the selected range to green.</p>
            <br/>
            <h3>Try it out</h3>
            <button @click="onSetColor">Set color</button>
          </div>
      </div>
    </div>
  </div>
</template>

<script>
export default {
  name: 'App',
  methods: {
    onSetColor () {
      window.Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange()
        range.format.fill.color = 'green'
        await context.sync()
      })
    }
  }
}
</script>

<style>
#content-header {
    background: #2a8dd4;
    color: #fff;
    position: absolute;
    top: 0;
    left: 0;
    width: 100%;
    height: 80px;
    overflow: hidden;
}

#content-main {
    background: #fff;
    position: fixed;
    top: 80px;
    left: 0;
    right: 0;
    bottom: 0;
    overflow: auto;
}

.padding {
    padding: 15px;
}
</style>