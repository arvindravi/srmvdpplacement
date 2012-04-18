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
    
    @students = Student.sort(:regno.desc)

    Spreadsheet.client_encoding = 'UTF-8'
    book = Spreadsheet::Workbook.new
    sheet1 = book.create_worksheet
    sheet2 = book.create_worksheet
    sheet3 = book.create_worksheet
    

    sheet1.row(0).concat %w{StudentName Registration_Number Branch DOB Gender Year_From Year_To Board_Tenth Mark_Tenth Month_Of_Passing_Tenth Year_Of_Passing_Tenth Board_Twelth Mark_Twelth Month_Of_Passing_Twelth Year_Of_Passing_Twelth Diploma_Name_Of_College Mark_Diploma Month_Of_Passsing_Diploma Year_Of_Passing_Diploma Break_Years Break_Reason Sem1_GPA Sem1_History_Of_Arrears Sem1_Standing_Arreas Sem2_GPA Sem2_History_Of_Arrears Sem2_Standing_Arreas Sem3_GPA Sem3_History_Of_Arrears Sem3_Standing_Arreas Sem4_GPA Sem4_History_Of_Arrears Sem4_Standing_Arreas Sem5_GPA Sem5_History_Of_Arrears Sem5_Standing_Arreas Sem6_GPA Sem6_History_Of_Arrears Sem6_Standing_Arreas Sem7_GPA Sem7_History_Of_Arrears Sem7_Standing_Arreas Sem8_GPA Sem8_History_Of_Arrears Sem8_Standing_Arreas At_The_End_Of_Course CGPA Total_Arrears Inplant_Training1_Organization_Name Inplant_Training1_Duration_From Inplant_Training1_Duration_To Inplant_Training1_Area_Of_Work Inplant_Training1_Total_Days Inplant_Training2_Organization_Name Inplant_Training2_Duration_From Inplant_Training2_Duration_To Inplant_Training2_Area_Of_Work Inplant_Training2_Total_Days Inplant_Training3_Organization_Name Inplant_Training3_Duration_From Inplant_Training3_Duration_To Inplant_Training3_Area_Of_Work Inplant_Training3_Total_Days Electives Mother_Tongue Languages_I_Can_Speak Languages_I_Can_Read Languages_I_Can_Understand Height Weight Nationality Passport_Number Passport_Valid_Upto Address_Of_Communication Phone1 Permanent_Address Phone2 Mobile Email_Address Father_Name Father_Occupation Father_Designation Father_Organisation_Address Father_Phone_L Father_Phone_M Father_Email Mother_Name Mother_Occupation Mother_Designation Mother_Organisation_Address Mother_Phone_L Mother_Phone_M Mother_Email Activities Stay Placement Plcement_No_Reason }
    
    sheet2.row(0).concat %w{StudentName Registration_Number Branch DOB Gender Year_From Year_To Board_Tenth Mark_Tenth Month_Of_Passing_Tenth Year_Of_Passing_Tenth Board_Twelth Mark_Twelth Month_Of_Passing_Twelth Year_Of_Passing_Twelth Diploma_Name_Of_College Mark_Diploma Month_Of_Passsing_Diploma Year_Of_Passing_Diploma Break_Years Break_Reason Sem1_GPA Sem1_History_Of_Arrears Sem1_Standing_Arreas Sem2_GPA Sem2_History_Of_Arrears Sem2_Standing_Arreas Sem3_GPA Sem3_History_Of_Arrears Sem3_Standing_Arreas Sem4_GPA Sem4_History_Of_Arrears Sem4_Standing_Arreas Sem5_GPA Sem5_History_Of_Arrears Sem5_Standing_Arreas Sem6_GPA Sem6_History_Of_Arrears Sem6_Standing_Arreas Sem7_GPA Sem7_History_Of_Arrears Sem7_Standing_Arreas Sem8_GPA Sem8_History_Of_Arrears Sem8_Standing_Arreas At_The_End_Of_Course CGPA Total_Arrears Inplant_Training1_Organization_Name Inplant_Training1_Duration_From Inplant_Training1_Duration_To Inplant_Training1_Area_Of_Work Inplant_Training1_Total_Days Inplant_Training2_Organization_Name Inplant_Training2_Duration_From Inplant_Training2_Duration_To Inplant_Training2_Area_Of_Work Inplant_Training2_Total_Days Inplant_Training3_Organization_Name Inplant_Training3_Duration_From Inplant_Training3_Duration_To Inplant_Training3_Area_Of_Work Inplant_Training3_Total_Days Electives Mother_Tongue Languages_I_Can_Speak Languages_I_Can_Read Languages_I_Can_Understand Height Weight Nationality Passport_Number Passport_Valid_Upto Address_Of_Communication Phone1 Permanent_Address Phone2 Mobile Email_Address Father_Name Father_Occupation Father_Designation Father_Organisation_Address Father_Phone_L Father_Phone_M Father_Email Mother_Name Mother_Occupation Mother_Designation Mother_Organisation_Address Mother_Phone_L Mother_Phone_M Mother_Email Activities Stay Placement Plcement_No_Reason }
    
    
    sheet3.row(0).concat %w{StudentName Registration_Number Branch DOB Gender Year_From Year_To Board_Tenth Mark_Tenth Month_Of_Passing_Tenth Year_Of_Passing_Tenth Board_Twelth Mark_Twelth Month_Of_Passing_Twelth Year_Of_Passing_Twelth Diploma_Name_Of_College Mark_Diploma Month_Of_Passsing_Diploma Year_Of_Passing_Diploma Break_Years Break_Reason Sem1_GPA Sem1_History_Of_Arrears Sem1_Standing_Arreas Sem2_GPA Sem2_History_Of_Arrears Sem2_Standing_Arreas Sem3_GPA Sem3_History_Of_Arrears Sem3_Standing_Arreas Sem4_GPA Sem4_History_Of_Arrears Sem4_Standing_Arreas Sem5_GPA Sem5_History_Of_Arrears Sem5_Standing_Arreas Sem6_GPA Sem6_History_Of_Arrears Sem6_Standing_Arreas Sem7_GPA Sem7_History_Of_Arrears Sem7_Standing_Arreas Sem8_GPA Sem8_History_Of_Arrears Sem8_Standing_Arreas At_The_End_Of_Course CGPA Total_Arrears Inplant_Training1_Organization_Name Inplant_Training1_Duration_From Inplant_Training1_Duration_To Inplant_Training1_Area_Of_Work Inplant_Training1_Total_Days Inplant_Training2_Organization_Name Inplant_Training2_Duration_From Inplant_Training2_Duration_To Inplant_Training2_Area_Of_Work Inplant_Training2_Total_Days Inplant_Training3_Organization_Name Inplant_Training3_Duration_From Inplant_Training3_Duration_To Inplant_Training3_Area_Of_Work Inplant_Training3_Total_Days Electives Mother_Tongue Languages_I_Can_Speak Languages_I_Can_Read Languages_I_Can_Understand Height Weight Nationality Passport_Number Passport_Valid_Upto Address_Of_Communication Phone1 Permanent_Address Phone2 Mobile Email_Address Father_Name Father_Occupation Father_Designation Father_Organisation_Address Father_Phone_L Father_Phone_M Father_Email Mother_Name Mother_Occupation Mother_Designation Mother_Organisation_Address Mother_Phone_L Mother_Phone_M Mother_Email Activities Stay Placement Plcement_No_Reason }
    
    
    i = 1 
    j = 1
    k = 1 
    @students.each do |s|
        
    
    if s.regno.to_i > 1040000000 && s.regno.to_i < 1041000000
      
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
      sheet1[i, 54] = s.imp2_d_f
      sheet1[i, 55] = s.imp2_d_t
      
      sheet1[i, 56] = s.imp2_w
      sheet1[i, 57] = s.imp2_dd
      sheet1[i, 58] = s.imp3_n
      sheet1[i, 59] = s.imp3_d_f
      sheet1[i, 60] = s.imp3_d_t
      sheet1[i, 61] = s.imp3_w
      sheet1[i, 62] = s.imp3_dd 
      sheet1[i, 63] = s.electives
      sheet1[i, 64] = s.m_t
      sheet1[i, 65] = s.lang_speak
      sheet1[i, 66] = s.lang_read
      sheet1[i, 67] = s.lang_und
      sheet1[i, 68] = s.height
      sheet1[i, 69] = s.weight
      sheet1[i, 70] = s.natl
      sheet1[i, 71] = s.ppn
      sheet1[i, 72] = s.pp_v 
      
      
      sheet1[i, 73] = s.address
      sheet1[i, 74] = s.ph1
      sheet1[i, 75] = s.add_p
      sheet1[i, 76] = s.ph2
      sheet1[i, 77] = s.mob
      sheet1[i, 78] = s.email
      sheet1[i, 79] = s.father_n
      sheet1[i, 80] = s.father_o
      sheet1[i, 81] = s.father_d
      sheet1[i, 82] = s.father_org 
      sheet1[i, 83] = s.father_ph
      sheet1[i, 84] = s.father_mob
      sheet1[i, 85] = s.father_email
      sheet1[i, 86] = s.mother_n
      sheet1[i, 87] = s.mother_o
      sheet1[i, 88] = s.mother_d
      sheet1[i, 89] = s.mother_org
      sheet1[i, 90] = s.mother_ph
      sheet1[i, 91] = s.mother_mob
      sheet1[i, 92] = s.mother_email 
      sheet1[i, 93] = s.activities
      sheet1[i, 94] = s.stay
      sheet1[i, 95] = s.pl
      sheet1[i, 96] = s.no_res
      
      i = i+ 1
      
    elsif s.regno.to_i > 1030000000 && s.regno.to_i < 1031000000
      
       sheet2[j, 0] = s.name
        sheet2[j, 1] = s.regno
        sheet2[j, 2] = s.branch
        sheet2[j, 3] = "#{s.d}/#{s.m}#{s.y}"
        sheet2[j, 4] = s.gender
        sheet2[j, 5] = s.a_y1
        sheet2[j, 6] = s.a_y2
        sheet2[j, 7] = s.board_tenth
        sheet2[j, 8] = s.mark_tenth 
        sheet2[j, 9] = s.monthofpassing_tenth
        sheet2[j, 10] = s.yearofpassing_tenth
        sheet2[j, 11] = s.board_twelth
        sheet2[j, 12] = s.mark_twelth
        sheet2[j, 13] = s.monthofpassing_twelth
        sheet2[j, 14] = s.yearofpassing_twelth
        sheet2[j, 15] = s.coll_diploma
        sheet2[j, 16] = s.diploma_mark
        sheet2[j, 17] = s.diploma_month
        sheet2[j, 18] = s.diploma_year 
        sheet2[j, 19] = s.break_y
        sheet2[j, 20] = s.break_r
        sheet2[j, 21] = s.sem1_gpa
        sheet2[j, 22] = s.sem1_a
        sheet2[j, 23] = s.sem1_sa
        sheet2[j, 24] = s.sem2_gpa
        sheet2[j, 25] = s.sem2_a
        sheet2[j, 26] = s.sem2_sa
        sheet2[j, 27] = s.sem3_gpa
        sheet2[j, 28] = s.sem3_a 
        sheet2[j, 29] = s.sem3_sa
        sheet2[j, 30] = s.sem4_gpa
        sheet2[j, 31] = s.sem4_a
        sheet2[j, 32] = s.sem4_sa
        sheet2[j, 33] = s.sem5_gpa
        sheet2[j, 34] = s.sem5_a
        sheet2[j, 35] = s.sem5_sa
        sheet2[j, 36] = s.sem6_gpa
        sheet2[j, 37] = s.sem6_a
        sheet2[j, 38] = s.sem6_sa 
        sheet2[j, 39] = s.sem7_gpa
        sheet2[j, 40] = s.sem7_a
        sheet2[j, 41] = s.sem7_sa
        sheet2[j, 42] = s.sem8_gpa
        sheet2[j, 43] = s.sem8_a
        sheet2[j, 44] = s.sem8_sa
        sheet2[j, 45] = s.eoc
        sheet2[j, 46] = s.cgpa
        sheet2[j, 47] = s.tot_a
        sheet2[j, 48] = s.imp1_n 
        sheet2[j, 49] = s.imp1_d_f
        sheet2[j, 50] = s.imp1_d_t
        sheet2[j, 51] = s.imp1_w
        sheet2[j, 52] = s.imp1_dd
        sheet2[j, 53] = s.imp2_n
        sheet2[j, 54] = s.imp2_d_f
        sheet2[j, 55] = s.imp2_d_t

        sheet2[j, 56] = s.imp2_w
        sheet2[j, 57] = s.imp2_dd
        sheet2[j, 58] = s.imp3_n
        sheet2[j, 59] = s.imp3_d_f
        sheet2[j, 60] = s.imp3_d_t
        sheet2[j, 61] = s.imp3_w
        sheet2[j, 62] = s.imp3_dd 
        sheet2[j, 63] = s.electives
        sheet2[j, 64] = s.m_t
        sheet2[j, 65] = s.lang_speak
        sheet2[j, 66] = s.lang_read
        sheet2[j, 67] = s.lang_und
        sheet2[j, 68] = s.height
        sheet2[j, 69] = s.weight
        sheet2[j, 70] = s.natl
        sheet2[j, 71] = s.ppn
        sheet2[j, 72] = s.pp_v 


        sheet2[j, 73] = s.address
        sheet2[j, 74] = s.ph1
        sheet2[j, 75] = s.add_p
        sheet2[j, 76] = s.ph2
        sheet2[j, 77] = s.mob
        sheet2[j, 78] = s.email
        sheet2[j, 79] = s.father_n
        sheet2[j, 80] = s.father_o
        sheet2[j, 81] = s.father_d
        sheet2[j, 82] = s.father_org 
        sheet2[j, 83] = s.father_ph
        sheet2[j, 84] = s.father_mob
        sheet2[j, 85] = s.father_email
        sheet2[j, 86] = s.mother_n
        sheet2[j, 87] = s.mother_o
        sheet2[j, 88] = s.mother_d
        sheet2[j, 89] = s.mother_org
        sheet2[j, 90] = s.mother_ph
        sheet2[j, 91] = s.mother_mob
        sheet2[j, 92] = s.mother_email 
        sheet2[j, 93] = s.activities
        sheet2[j, 94] = s.stay
        sheet2[j, 95] = s.pl
        sheet2[j, 96] = s.no_res
        
        j=j+1
      
      
    else
      
       sheet3[k, 0] = s.name
        sheet3[k, 1] = s.regno
        sheet3[k, 2] = s.branch
        sheet3[k, 3] = "#{s.d}/#{s.m}#{s.y}"
        sheet3[k, 4] = s.gender
        sheet3[k, 5] = s.a_y1
        sheet3[k, 6] = s.a_y2
        sheet3[k, 7] = s.board_tenth
        sheet3[k, 8] = s.mark_tenth 
        sheet3[k, 9] = s.monthofpassing_tenth
        sheet3[k, 10] = s.yearofpassing_tenth
        sheet3[k, 11] = s.board_twelth
        sheet3[k, 12] = s.mark_twelth
        sheet3[k, 13] = s.monthofpassing_twelth
        sheet3[k, 14] = s.yearofpassing_twelth
        sheet3[k, 15] = s.coll_diploma
        sheet3[k, 16] = s.diploma_mark
        sheet3[k, 17] = s.diploma_month
        sheet3[k, 18] = s.diploma_year 
        sheet3[k, 19] = s.break_y
        sheet3[k, 20] = s.break_r
        sheet3[k, 21] = s.sem1_gpa
        sheet3[k, 22] = s.sem1_a
        sheet3[k, 23] = s.sem1_sa
        sheet3[k, 24] = s.sem2_gpa
        sheet3[k, 25] = s.sem2_a
        sheet3[k, 26] = s.sem2_sa
        sheet3[k, 27] = s.sem3_gpa
        sheet3[k, 28] = s.sem3_a 
        sheet3[k, 29] = s.sem3_sa
        sheet3[k, 30] = s.sem4_gpa
        sheet3[k, 31] = s.sem4_a
        sheet3[k, 32] = s.sem4_sa
        sheet3[k, 33] = s.sem5_gpa
        sheet3[k, 34] = s.sem5_a
        sheet3[k, 35] = s.sem5_sa
        sheet3[k, 36] = s.sem6_gpa
        sheet3[k, 37] = s.sem6_a
        sheet3[k, 38] = s.sem6_sa 
        sheet3[k, 39] = s.sem7_gpa
        sheet3[k, 40] = s.sem7_a
        sheet3[k, 41] = s.sem7_sa
        sheet3[k, 42] = s.sem8_gpa
        sheet3[k, 43] = s.sem8_a
        sheet3[k, 44] = s.sem8_sa
        sheet3[k, 45] = s.eoc
        sheet3[k, 46] = s.cgpa
        sheet3[k, 47] = s.tot_a
        sheet3[k, 48] = s.imp1_n 
        sheet3[k, 49] = s.imp1_d_f
        sheet3[k, 50] = s.imp1_d_t
        sheet3[k, 51] = s.imp1_w
        sheet3[k, 52] = s.imp1_dd
        sheet3[k, 53] = s.imp2_n
        sheet3[k, 54] = s.imp2_d_f
        sheet3[k, 55] = s.imp2_d_t

        sheet3[k, 56] = s.imp2_w
        sheet3[k, 57] = s.imp2_dd
        sheet3[k, 58] = s.imp3_n
        sheet3[k, 59] = s.imp3_d_f
        sheet3[k, 60] = s.imp3_d_t
        sheet3[k, 61] = s.imp3_w
        sheet3[k, 62] = s.imp3_dd 
        sheet3[k, 63] = s.electives
        sheet3[k, 64] = s.m_t
        sheet3[k, 65] = s.lang_speak
        sheet3[k, 66] = s.lang_read
        sheet3[k, 67] = s.lang_und
        sheet3[k, 68] = s.height
        sheet3[k, 69] = s.weight
        sheet3[k, 70] = s.natl
        sheet3[k, 71] = s.ppn
        sheet3[k, 72] = s.pp_v 


        sheet3[k, 73] = s.address
        sheet3[k, 74] = s.ph1
        sheet3[k, 75] = s.add_p
        sheet3[k, 76] = s.ph2
        sheet3[k, 77] = s.mob
        sheet3[k, 78] = s.email
        sheet3[k, 79] = s.father_n
        sheet3[k, 80] = s.father_o
        sheet3[k, 81] = s.father_d
        sheet3[k, 82] = s.father_org 
        sheet3[k, 83] = s.father_ph
        sheet3[k, 84] = s.father_mob
        sheet3[k, 85] = s.father_email
        sheet3[k, 86] = s.mother_n
        sheet3[k, 87] = s.mother_o
        sheet3[k, 88] = s.mother_d
        sheet3[k, 89] = s.mother_org
        sheet3[k, 90] = s.mother_ph
        sheet3[k, 91] = s.mother_mob
        sheet3[k, 92] = s.mother_email 
        sheet3[k, 93] = s.activities
        sheet3[k, 94] = s.stay
        sheet3[k, 95] = s.pl
        sheet3[k, 96] = s.no_res
        
        k=k+1
      
    end
          
      


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

