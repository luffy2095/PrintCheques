#!/usr/bin/ruby
#require "pathname"
require 'spreadsheet'
require 'prawn'
require "prawn/measurement_extensions"
require 'rupees'
##Please change the format of your cheque using following constants 
#CONSTANTS
Y_value=87                #Beginning of output of date on first cheque
S1=11				#Space between date and payee
S2=20				#Space between date and amount_words
S5=28				#Space between date and amount_numbers
S3=60				#Space between date and not exceeding
S4=37				#Space between date and Signature line
Payee=Y_value - S1
Amount=Y_value - S2
Exceed=Y_value - S3
Infi=Y_value - S4
Amt_num=Y_value - S5
X1=161				#Horizontal spacing date
X2=21				#Horizontal spacing payee
X3=29				#Horizontal spacing amount
X4=163				#Horizontal spacing amount numbers
X5=72				#Horizontal spacing not exceeding
X6=139				#Horizontal spacing Signature line
Height=25			#Height of each field
Width_short=80			#Width of short fields
Width_long=120	
Spacing=5		#Width of long fields
####################################################################
#Begin block
BEGIN{

	puts "Please Enter Absolute path of Excel file,if both are not in same Directory.\n"
	puts "Else just enter name of file. Extention necessary.\n"
	puts "Please check that the file format is 'xls'. Convert if necessary"
	print "Parsing..."
			
}
END{	puts "\n \nThe valid cheques have been printed."
	puts "Please check 'errors.xls' for errors in cheques that have been omitted \n"
}
#################
class Automation



####Main arrays
@@name=Array.new
@@amount_num=Array.new
@@amount_words=Array.new
@@date=Array.new
@@limit=Array.new
@@NoParsed=0
####Arrays to store errors
@@e_name=Array.new
@@e_amount_num=Array.new
@@e_amount_words=Array.new
@@e_date=Array.new
@@e_limit=Array.new
@@e_description=Array.new
@@NoOfErrors=0

###############################METHODS###############################

################Accessor method########################
def getNoOfRows
	@@NoParsed
end
######################################################
def errorHandling(tempName,tempAmtNum,tempAmtWords,tempDate,tempLimit,description)
i=@@NoOfErrors
@@e_name[i]=tempName
@@e_amount_num[i]=tempAmtNum
@@e_amount_words[i]=tempAmtWords
@@e_date[i]=tempDate
@@e_limit[i]=tempLimit
@@e_description[i]=description
@@NoOfErrors+=1
end
##################################

def fileParsing(f)		##this method parses the excel file

	book1 = Spreadsheet.open(f,'rb')
	book1.worksheets.each do |sheet|           
		iArrayNo=0		##denotes position in arrays
		sheet.each 1 do |row|
		tempName=row[0]
		tempAmtNum=row[1]
		tempAmtWords=row[2]
		tempDate=row[3]
		tempLimit=row[4]
		#if tempLimit!=nil and tempAmtNum>tempLimit
		#	errorHandling(tempName,tempAmtNum,tempAmtWords,tempDate,tempLimit,"Amount exceeded limit")
		#	next
		#end	
	   	@@name[iArrayNo]=row[0]
		@@amount_num[iArrayNo]=row[1]
	   	if @@amount_num[iArrayNo]!=nil
		@@amount_num[iArrayNo]=@@amount_num[iArrayNo].to_i
		end
	   	@@amount_words[iArrayNo]=row[2]
		
	   	@@date[iArrayNo]=row[3]
		if @@date[iArrayNo]!=nil
		@@date[iArrayNo]=@@date[iArrayNo].to_s
		end
		
		@@limit[iArrayNo]=row[4]
		if @@limit[iArrayNo]!=nil
	   	@@limit[iArrayNo]=@@limit[iArrayNo].to_i
		end
		
	   	@@NoParsed+=1
	   	iArrayNo+=1
		end  
		
	end

end					##fileParsing method ends here

###################################################

#def printing			##temp method
#	puts @@name
#	puts "\n"
#	puts @@e_name
#	puts @@NoOfErrors
#end				##printing ends here

###################################################33

def make_pdf(n_of_checks)     #Method to create pdf
Prawn::Document.generate("Printpdf.pdf",
			:print_scaling => :none,
			:page_size => [210.mm, 94.mm],
			:margin => 0.mm) do
	

	#Following 2 lines just for trial, to be deleted
	#index=1
	#n_of_checks=1
	#Above 2 lines for trial	
	
	

	
	for i in 0..n_of_checks      #Beginning of For statement

	text_box "#{@@date[i]}",
	:at => [(::X1).mm, (Y_value).mm],
	:character_spacing => 3.mm,
	:height => ::Height.mm,
	:width => ::Width_short.mm
	
	text_box "#{@@name[i]}",
	:at => [(::X2).mm, (Payee).mm],
	:height => ::Height.mm,
	:width => ::Width_long.mm
	
	if @@amount_words[i]!=nil
	text_box "#{@@amount_words[i]} only",
	:at => [(::X3).mm, (Amount).mm],
	:height => ::Height.mm,
	:leading => ::Spacing.mm,
	:width => ::Width_long.mm
	end
	
	if @@amount_words[i]==nil && @@amount_num[i]!=nil                        #using rupees gem
	text_box "#{@@amount_num[i].rupees} only",
	:at => [(::X3).mm, (Amount).mm],
	:height => ::Height.mm,
	:leading => ::Spacing.mm,
	:width => ::Width_long.mm
	end
	
	if @@amount_num[i]!=nil
	text_box "**#{@@amount_num[i]}/-",
	:at => [(::X4).mm, (Amt_num).mm],
	:height => ::Height.mm,
	:width => ::Width_short.mm
	end
	
	if @@limit[i]!=nil	
	text_box "Not over Rs.#{@@limit[i]}/-",
	:at => [(::X5).mm, (Exceed).mm],
	:height => ::Height.mm,
	:width => ::Width_short.mm
	end

	#text_box "For NSI Infinium Global Pvt. Ltd.",#FIXED
	#:at => [(::X6).mm, (Infi).mm],
	#:height => ::Height.mm,
	#:width => ::Width_short.mm
	
	if i!=n_of_checks
	start_new_page(:size => [210.mm, 94.mm],:margin => 0.mm)
	end
	end	#end of for statement
	
end          #end of the generated pdf

#Now to print the pdf that was created.
	#system("lpr", "Printpdf.pdf")	


end      #end of make_pdf method

#####################################################


def errorPrinting		##handles 'errors.xls' file
	count=@@NoOfErrors
	#puts count
	i=2
	book=Spreadsheet::Workbook.new
	sheet2=book.create_worksheet
	sheet2.column(0).width=20
	sheet2.column(1).width=20
	sheet2.column(2).width=40
	sheet2.column(3).width=20
	sheet2.column(4).width=20
	sheet2.column(5).width=60
	format = Spreadsheet::Format.new    :color => :blue,
                                   		:weight => :bold,
                                   		:size => 10,
                                   		:align => :center
        format1 = Spreadsheet::Format.new      :align => :center
                                   		
	sheet2.row(0).default_format = format
	sheet2.column(1).default_format=format1
	sheet2.column(3).default_format=format1
	sheet2.column(4).default_format=format1
	sheet2.column(5).default_format=format1
	sheet2.name='error'				##
	sheet2[0,0]='NAME OF PAYEE'				##
	sheet2[0,1]='AMOUNT IN FIGURES'			##
	sheet2[0,2]='AMOUNT IN WORDS'			##   for heading of errors.xls
	sheet2[0,3]='DATE'				##
	sheet2[0,4]='NOT EXCEEDING Rs'				##
	sheet2[0,5]='DESCRIPTION'			##
	
	while count>0
	sheet2[i,0]=@@e_name[i-2]
	#puts @@e_name
	sheet2[i,1]=@@e_amount_num[i-2]
	sheet2[i,2]=@@e_amount_words[i-2]
	sheet2[i,3]=@@e_date[i-2]
	sheet2[i,4]=@@e_limit[i-2]
	sheet2[i,5]=@@e_description[i-2]
	i+=1
	count-=1
	end
	#puts @@e_name
	book.write 'errors.xls'
	
end			## errorPrinting ends here


end 				##class end here

Object creation
begin                     	##begin of a code snippet
pathName=gets.chomp		##input from user (filename) 
	if pathName==""         ##check if fileName entered
		raise           ##If file name not entered raise an exception
	end
rescue				##executed when exception raised
system "clear"
puts "please enter file name"   
retry				##retry from begin block
end 				##end of begin block
######################################################################

a=Automation.new		##object
a.fileParsing("cheque.xls")
noOfChecks=a.getNoOfRows
a.make_pdf(noOfChecks-1)
a.errorPrinting




