define(['jquery', 'use!backbone', 'text!app/templates/base_template.html',], function($, Backbone, html_template){
   window.BaseView = Backbone.View.extend({
      
       
       initialize: function() {
            this.template = _.template(html_template);
            this.render();
       },
       
       render: function(){
            $("#container").html(this.template);
       }
       
       

       
   });
   
   
   
});