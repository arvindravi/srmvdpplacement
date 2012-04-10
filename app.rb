require 'rubygems'
require 'sinatra'
require 'mongo'
require 'mongo_mapper'

require 'spreadsheet'
#require 'rack-flash'
require 'will_paginate'
require 'will_paginate-bootstrap'
require 'mongomapper_search'
require_relative 'models'

uri = "mongodb://heroku_app3705114:80e932ba937qf19bk9d0t65hf8@ds031747.mongolab.com:31747/heroku_app3705114"
#uri =  "mongodb://heroku_app3666321:9nvpgjk8aglthg3qfm6j8nof0o@ds031777.mongolab.com:31777/heroku_app3666321"
#uri = "mongodb://localhost:27017"
MongoMapper.connection = Mongo::Connection.from_uri( uri )
MongoMapper.database = 'heroku_app3705114'
#MongoMapper.database = 'heroku_app3666321'

enable :sessions

get '/' do
  erb :index
end

#get '/badmarks' do 

#  @students = Student.sort(:regno)
  
#  erb :badmarks
  
#end

get '/ece/left' do
  
  @studentsleft = Array.new
  
  
  
  i = 1040940001
  
  296.times do 
    if !Student.find_by_regno(i.to_s)
      @studentsleft.push(i)
    end
    i += 1
  end
  @dept = "ECE"
  erb :showleft
  
end

get '/cse/left' do
  
  @studentsleft = Array.new
  
  
  i = 1030940001
  
  73.times do 
    if !Student.find_by_regno(i.to_s)
      @studentsleft.push(i)
    end
    i += 1
  end
  
  @dept = "CSE"
  erb :showleft
  
end

get '/mech/left' do
  
  @studentsleft = Array.new
  
  
  i = 1020940001
  
  210.times do 
    if !Student.find_by_regno(i.to_s)
      @studentsleft.push(i)
    end
    i += 1
  end
  
  @dept = "Mechanical"
  erb :showleft
  
end


get '/unblock' do 
  
  students = Student.all
  
  students.each do |s|
    s.attempt = 1
    s.save

  end
  
  "Done!"
  
end

post '/' do 
  
  
  @app = App.find_by_aident("stf0001")
  
  if @app.open == 1 
    m = /10(2|3|4)09(4|1)0\d\d\d/.match(params[:student][:regno])
      if m.nil?
        session[:error] = "<strong>Invalid Registration Number!</strong>"
        redirect '/'

      else 

        student = Student.find_by_regno(params[:student][:regno])

        if student && student.attempt == 1
          redirect "/ug/edit/#{student.id}"
        elsif student && student.attempt == 2
           redirect "/ug/readonly/#{student.id}"
        else
          redirect "/ug/#{params[:student][:regno]}"
        end
      end
      
    else
      
      erb :closed
  end
    
end

get '/ug/readonly/:id' do |sid|
  @s = Student.find_by_id(sid)
  erb :readonly
end

get '/ug/:regno' do |r|
  @regno = r
  erb :ug
end

get '/ug/edit/:id' do |id|
  @s = Student.find_by_id(id)
  erb :edit
end

put '/ug/:id/save' do |id|
  @s = Student.find_by_id(params[:id])
  pa = @s.attempt
  @s.attempt = pa + 1
  if @s.update_attributes(params[:student])
    session[:notice] = "Your information was updated successfully! Thanks!<br />
    <strong> All the best for your placement!</strong>"
    redirect '/'
  else
    session[:error] = "There was an error processing your request.. Please contact the administrator immediately!"
    redirect '/'
  end
end



post '/save' do 
  params[:student][:attempt] = 1
  student = Student.new(params[:student])
  #student.attempt = 1
  if student.save 
    session[:notice] = "Your information was saved successfully! Thanks!<br />
    <strong> All the best for your placement!</strong>"
    redirect '/'
  else
    session[:error] = "There was an error processing your request.. Please contact the administrator immediately!"
    redirect '/'
  end
end








get '/admin' do 

  erb :admin
  
end

post '/admin/login' do

if params[:user][:username] == "admin" and params[:user][:password] == "13132727" 
  session[:admin] = "Admin"
  redirect '/admin/home'
else
  redirect '/admin'
end
end

get '/admin/home' do
  if session[:admin] 
    @app = App.find_by_aident("stf0001")
    @sitestatus = @app.open
    @recordcount = Student.count     
    @students = Student.paginate(:page => params[:page], :per_page => 10)
    erb :adminhome
    
  else 
    redirect '/admin'
  end
  
end

get '/site/close' do

  if session[:admin]
    @app = App.find_by_aident("stf0001")
    @app.open = 0
    if @app.save
      session[:admin_notice] = "Website closed students cannot enter data anymore!"
      redirect "/admin/home"
    else
      session[:admin_warn] = "Could not shut down website please contact the administrator!"
      redirect "/admin/home"
    end
    redirect "/admin"
  end
  
  
end

get '/site/open' do

  if session[:admin]
    @app = App.find_by_aident("stf0001")
    @app.open = 1
    if @app.save
      session[:admin_notice] = "Website opened successfully!"
      redirect "/admin/home"
    else
      session[:admin_warn] = "Could not shut open website please contact the administrator!"
      redirect "/admin/home"
    end
    redirect "/admin"
  end
  
  
end

get '/admin/search' do

  @s = Student.find_by_regno(params[:query])
  erb :adminsearch
  
end

get '/info/:id' do |id|
  
    @s = Student.find_by_id(id)
    erb :fulldetail
  
end

get '/student/unblock/:id' do |id|
  @s = Student.find_by_id(id)
  
  if @s.update_attributes(:attempt => 1)
    session[:admin_notice] = "#{@s.regno} was unblocked sucessfully!"
    redirect "/admin/home"
  else
    session[:admin_warn] = "#{@s.regno} could not be unblocked!"
    redirect "/admin/home"
  end
end

get '/student/del/:id' do |id|
  @s = Student.find_by_id(id)
  @s.destroy if !@s.nil? 
  session[:admin_notice] = "#{@s.regno} was deleted sucessfully!"
  redirect "/admin/home"
end


get '/admin/logout' do
  session[:admin] = nil
  "Logout!"
end



get '/make' do
  if session[:admin] 
    
    @students = Student.all

    Spreadsheet.client_encoding = 'UTF-8'
    book = Spreadsheet::Workbook.new
    sheet1 = book.create_worksheet

    sheet1.row(0).concat %w{StudentName Registration_Number Branch DOB Gender Year_From Year_To Board_Tenth Mark_Tenth Month_Of_Passing_Tenth Year_Of_Passing_Tenth Board_Twelth Mark_Twelth Month_Of_Passing_Twelth Year_Of_Passing_Twelth Diploma_Name_Of_College Mark_Diploma Month_Of_Passsing_Diploma Year_Of_Passing_Diploma Break_Years Break_Reason Sem1_GPA Sem1_History_Of_Arrears Sem1_Standing_Arreas Sem2_GPA Sem2_History_Of_Arrears Sem2_Standing_Arreas Sem3_GPA Sem3_History_Of_Arrears Sem3_Standing_Arreas Sem4_GPA Sem4_History_Of_Arrears Sem4_Standing_Arreas Sem5_GPA Sem5_History_Of_Arrears Sem5_Standing_Arreas Sem6_GPA Sem6_History_Of_Arrears Sem6_Standing_Arreas Sem7_GPA Sem7_History_Of_Arrears Sem7_Standing_Arreas Sem8_GPA Sem8_History_Of_Arrears Sem8_Standing_Arreas At_The_End_Of_Course CGPA Total_Arrears Inplant_Training1_Organization_Name Inplant_Training1_Duration_From Inplant_Training1_Duration_To Inplant_Training1_Area_Of_Work Inplant_Training1_Total_Days Inplant_Training2_Organization_Name Inplant_Training2_Duration_From Inplant_Training2_Duration_To Inplant_Training2_Area_Of_Work Inplant_Training2_Total_Days Inplant_Training3_Organization_Name Inplant_Training3_Duration_From Inplant_Training3_Duration_To Inplant_Training3_Area_Of_Work Inplant_Training3_Total_Days Electives Mother_Tongue Languages_I_Can_Speak Languages_I_Can_Read Languages_I_Can_Understand Height Weight Nationality Passport_Number Passport_Valid_Upto Address_Of_Communication Phone1 Permanent_Address Phone2 Mobile Email_Address Father_Name Father_Occupation Father_Designation Father_Organisation_Address Father_Phone_L Father_Phone_M Father_Email Mother_Name Mother_Occupation Mother_Designation Mother_Organisation_Address Mother_Phone_L Mother_Phone_M Mother_Email Activities Stay Placement Plcement_No_Reason }
    i = 1; 
    @students.each do |s|
        
    
      
      sheet1[i, 0] = s.name
      sheet1[i, 1] = s.regno
      sheet1[i, 2] = s.branch
      sheet1[i, 3] = "#{s.d}/#{s.m}#{s.y}"
      sheet1[i, 4] = s.gender
      sheet1[i, 5] = s.a_y1
      sheet1[i, 6] = s.a_y2
      sheet1[i, 7] = s.board_tenth
      sheet1[i, 8] = s.mark_tenth 
      sheet1[i, 9] = s.monthofpassing_tenth
      sheet1[i, 10] = s.yearofpassing_tenth
      sheet1[i, 11] = s.board_twelth
      sheet1[i, 12] = s.mark_twelth
      sheet1[i, 13] = s.monthofpassing_twelth
      sheet1[i, 14] = s.yearofpassing_twelth
      sheet1[i, 15] = s.coll_diploma
      sheet1[i, 16] = s.diploma_mark
      sheet1[i, 17] = s.diploma_month
      sheet1[i, 18] = s.diploma_year 
      sheet1[i, 19] = s.break_y
      sheet1[i, 20] = s.break_r
      sheet1[i, 21] = s.sem1_gpa
      sheet1[i, 22] = s.sem1_a
      sheet1[i, 23] = s.sem1_sa
      sheet1[i, 24] = s.sem2_gpa
      sheet1[i, 25] = s.sem2_a
      sheet1[i, 26] = s.sem2_sa
      sheet1[i, 27] = s.sem3_gpa
      sheet1[i, 28] = s.sem3_a 
      sheet1[i, 29] = s.sem3_sa
      sheet1[i, 30] = s.sem4_gpa
      sheet1[i, 31] = s.sem4_a
      sheet1[i, 32] = s.sem4_sa
      sheet1[i, 33] = s.sem5_gpa
      sheet1[i, 34] = s.sem5_a
      sheet1[i, 35] = s.sem5_sa
      sheet1[i, 36] = s.sem6_gpa
      sheet1[i, 37] = s.sem6_a
      sheet1[i, 38] = s.sem6_sa 
      sheet1[i, 39] = s.sem7_gpa
      sheet1[i, 40] = s.sem7_a
      sheet1[i, 41] = s.sem7_sa
      sheet1[i, 42] = s.sem8_gpa
      sheet1[i, 43] = s.sem8_a
      sheet1[i, 44] = s.sem8_sa
      sheet1[i, 45] = s.eoc
      sheet1[i, 46] = s.cgpa
      sheet1[i, 47] = s.tot_a
      sheet1[i, 48] = s.imp1_n 
      sheet1[i, 49] = s.imp1_d_f
      sheet1[i, 50] = s.imp1_d_t
      sheet1[i, 51] = s.imp1_w
      sheet1[i, 52] = s.imp1_dd
      sheet1[i, 53] = s.imp2_n
      sheet1[i, 54] = s.imp2_w
      sheet1[i, 55] = s.imp2_dd
      sheet1[i, 56] = s.imp3_n
      sheet1[i, 57] = s.imp3_w
      sheet1[i, 58] = s.imp3_dd 
      sheet1[i, 59] = s.electives
      sheet1[i, 60] = s.m_t
      sheet1[i, 61] = s.lang_speak
      sheet1[i, 62] = s.lang_read
      sheet1[i, 63] = s.lang_und
      sheet1[i, 64] = s.height
      sheet1[i, 65] = s.weight
      sheet1[i, 66] = s.natl
      sheet1[i, 67] = s.ppn
      sheet1[i, 68] = s.pp_v 
      sheet1[i, 69] = s.address
      sheet1[i, 70] = s.ph1
      sheet1[i, 71] = s.add_p
      sheet1[i, 72] = s.ph2
      sheet1[i, 73] = s.mob
      sheet1[i, 74] = s.email
      sheet1[i, 75] = s.father_n
      sheet1[i, 76] = s.father_o
      sheet1[i, 77] = s.father_d
      sheet1[i, 78] = s.father_org 
      sheet1[i, 79] = s.father_ph
      sheet1[i, 80] = s.father_mob
      sheet1[i, 81] = s.father_email
      sheet1[i, 82] = s.mother_n
      sheet1[i, 83] = s.mother_o
      sheet1[i, 84] = s.mother_d
      sheet1[i, 85] = s.mother_org
      sheet1[i, 86] = s.mother_ph
      sheet1[i, 87] = s.mother_mob
      sheet1[i, 88] = s.mother_email 
      sheet1[i, 89] = s.activities
      sheet1[i, 90] = s.stay
      sheet1[i, 91] = s.pl
      sheet1[i, 92] = s.no_res
          
      
      i = i+1;

    end


        if book.write 'tmp/ugstudentlist.xls' 
          file = 'tmp/ugstudentlist.xls'
          send_file(file, :disposition => 'attachment', :filename => File.basename(file))

        else
           file = 'tmp/ugstudentlist.xls'
          send_file(file, :disposition => 'attachment', :filename => File.basename(file))

        end
        
      else 
        
        redirect '/admin'
    
  end
  
end

