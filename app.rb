require 'rubygems'
require 'sinatra'
require 'mongo'
require 'mongo_mapper'
require_relative 'models'
require 'spreadsheet'

uri =  "mongodb://heroku_app3666321:9nvpgjk8aglthg3qfm6j8nof0o@ds031777.mongolab.com:31777/heroku_app3666321"
#uri = "mongodb://localhost:27017"
MongoMapper.connection = Mongo::Connection.from_uri( uri )
MongoMapper.database = 'heroku_app3666321'


get '/' do
  File.read("index.html")
end

post '/' do 
   m = /103|2|104\d\d[0-9]\d/.match(params[:student][:reg_no])
    if m.nil?
      redirect '/'
    else 
      redirect '/home'
    end
end

get '/home' do
  File.read("home.html")
end

post '/save' do 
  student = Student.new(params[:student])
  if student.save 
    redirect '/list'
  end
end


get '/list' do
  @students = Student.all
  erb :list
end

get '/filter/:mark' do |mark|
  limit = (mark.to_i + 10).to_s
  mark_cgpa = mark.to_i / 10
  limit_cgpa = limit.to_i / 10
  
  @students = Student.all
  
  Spreadsheet.client_encoding = 'UTF-8'
  book = Spreadsheet::Workbook.new
  sheet1 = book.create_worksheet
  
  sheet1.row(0).concat %w{Registration_Number Student_Name Branch Department Campus Gender DOB 10th_Marks 10th_Passing 10th_Board 12th_Marks 12th_Passing 12th_Board CGPA}
  
  i = 1;
  @students.each do |s|
    
  if(((s.mark_tenth.to_i >= mark.to_i) && (s.mark_tenth.to_i <= limit.to_i)) && ((s.mark_twelth.to_i >= mark.to_i) && (s.mark_twelth.to_i <= limit.to_i)) && ((s.cgpa.to_i >= mark_cgpa) && (s.cgpa.to_i <= limit_cgpa))) 
    
    sheet1[i,0] = s.regno
    sheet1[i,1] = s.name
    sheet1[i, 2] = "BTech"
    sheet1[i, 3] = s.branch
    sheet1[i, 4] = "SRM Vadapalani"
    sheet1[i, 5] = s.gender
    sheet1[i, 6] = "#{s.d}/#{s.m}/#{s.y}"
    sheet1[i, 7] = s.mark_tenth
    sheet1[i, 8] = "#{s.monthofpassing_tenth} #{s.yearofpassing_tenth}"
    sheet1[i, 9] = s.board_tenth
    sheet1[i, 10] = s.mark_twelth
    sheet1[i, 11] = "#{ s.yearofpassing_twelth } #{ s.monthofpassing_twelth}"
    sheet1[i, 12] = s.board_twelth
     sheet1[i, 13] = s.cgpa
  
  end
    
    i = i+1;
    
  end
  
  book.write "test#{mark}.xls"
  #@students = Student.where(:mark_tenth => {:$gt => mark, :$lt => limi})
  
  #erb :filter
end

get '/arrear'  do

  @students = Student.all
  
  Spreadsheet.client_encoding = 'UTF-8'
  book = Spreadsheet::Workbook.new
  sheet1 = book.create_worksheet
  
  sheet1.row(0).concat %w{Registration_Number Student_Name Branch Department Campus Gender DOB 10th_Marks 10th_Passing 10th_Board 12th_Marks 12th_Passing 12th_Board CGPA}
  
  i = 1;
  @students.each do |s|
    
    
    if(s.tot_a.to_i > 0)
  
    
    sheet1[i,0] = s.regno
    sheet1[i,1] = s.name
    sheet1[i, 2] = "BTech"
    sheet1[i, 3] = s.branch
    sheet1[i, 4] = "SRM Vadapalani"
    sheet1[i, 5] = s.gender
    sheet1[i, 6] = "#{s.d}/#{s.m}/#{s.y}"
    sheet1[i, 7] = s.mark_tenth
    sheet1[i, 8] = "#{s.monthofpassing_tenth} #{s.yearofpassing_tenth}"
    sheet1[i, 9] = s.board_tenth
    sheet1[i, 10] = s.mark_twelth
    sheet1[i, 11] = "#{ s.yearofpassing_twelth } #{ s.monthofpassing_twelth}"
    sheet1[i, 12] = s.board_twelth
     sheet1[i, 13] = s.cgpa
  
end
    
    i = i+1;
    
  end
  
  book.write "arearlist.xls"
  

end

get '/make' do
  @students = Student.all
  
  Spreadsheet.client_encoding = 'UTF-8'
  book = Spreadsheet::Workbook.new
  sheet1 = book.create_worksheet
  
  sheet1.row(0).concat %w{Registration_Number Student_Name Branch Department Campus Gender DOB 10th_Marks 10th_Passing 10th_Board 12th_Marks 12th_Passing 12th_Board}
  i = 1; j = 0;
  @students.each do |s|
    
  if(s.mark_tenth.to_i > 70 ) 
    
    sheet1[i,0] = s.regno
    sheet1[i,1] = s.name
    sheet1[i, 2] = "BTech"
    sheet1[i, 3] = s.branch
    sheet1[i, 4] = "SRM Vadapalani"
    sheet1[i, 5] = s.gender
    sheet1[i, 6] = "#{s.d}/#{s.m}/#{s.y}"
    sheet1[i, 7] = s.mark_tenth
    sheet1[i, 8] = "#{s.monthofpassing_tenth} #{s.yearofpassing_tenth}"
    sheet1[i, 9] = s.board_tenth
    sheet1[i, 10] = s.mark_twelth
    sheet1[i, 11] = "#{ s.yearofpassing_twelth } #{ s.monthofpassing_twelth}"
    sheet1[i, 12] = s.board_twelth
  
  end
    
    i = i+1;
    
  end
  
  
  book.write "testlist.xls"
  
end

