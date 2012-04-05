require(['jquery', 'use!backbone', 'app/routes/approuter'], function($, Backbone){
    
           $(document).ready(function () {
              
       	    var appRouter = new AppRouter();
       	    Backbone.history.start();

       	  });
    
   
});