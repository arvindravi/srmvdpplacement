define(['use!backbone', 'app/views/base_view'], function(Backbone){
    window.AppRouter = Backbone.Router.extend({
            routes: {
                "": "defaultRoute"
            },

            defaultRoute: function( ){
               var baseView = new BaseView();
            }
        
        });
});