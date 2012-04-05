// Set the require.js configuration for your application.
require.config({
  // Initialize the application with the main application file
  deps: ["App"],

  paths: {
    // JavaScript folders
    libs: "lib/",
    //plugins: "../assets/js/plugins",

    // Libraries
    jquery: "lib/jquery",
    underscore: "lib/underscore",
    backbone: "lib/backbone",

    // Shim Plugin
    use: "plugins/use",
    //Text Plugin
    text: "plugins/text"
  },

  use: {
    backbone: {
      deps: ["use!underscore", "jquery"],
      attach: "Backbone"
    },

    underscore: {
      attach: "_"
    }
  }
});
