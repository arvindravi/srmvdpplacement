class Student 
include MongoMapper::Document


  
end

class App
include MongoMapper::Document

  key :aident, String
  key :open, Integer 
  
end
