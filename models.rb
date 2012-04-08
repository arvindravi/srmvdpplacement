class Student 
include MongoMapper::Document


  key :regno, String, :unique => true


end

class App
include MongoMapper::Document

  key :aident, String
  key :open, Integer 
  
end
