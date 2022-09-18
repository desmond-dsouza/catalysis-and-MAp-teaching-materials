COURSE_DIR = Dir.pwd
SOURCE = File.join COURSE_DIR, "source"
ZIPPED = File.join COURSE_DIR, "5-day-course.zip"

PRODUCTION = {
	:course_handout => {
		:front => ['00-cover'],
		:body => %w{01-introduction 02-method-overview 03-basic-modeling-skills 04-integrated-modeling-skills 05-business-architecture 06-black-box-architecture 07-white-box-architecture 08-technical-architecture 09-deployment-architecture 10-models-and-code 92-references},
		:back => ['new_case_study\\exercises_solutions.doc']
		},
	:exercise_handout => {
		:body => ['new_case_study\\exercises.doc']
	}
}

# TODO: create multiple FinePrint printers with specific settings

task :default do
  puts "Running rakefile.rb"
	FileUtils.cd SOURCE
	PRODUCTION.each {|outfile, spec|
		puts outfile
		[:front, :body, :back].each {|key|
			next if spec[key].nil?
			spec[key][0..2].each {|f|
				file = File.join(SOURCE, f).gsub('/','\\')
				puts "   " + file + "\n"
				fine_print file
			}
		}
	}
end

require 'win32ole'

module PPT; end
Ppt = WIN32OLE.new 'powerpoint.application'; Ppt.Visible = true
WIN32OLE.const_load(Ppt, PPT)

# module WORD; end
Word = WIN32OLE.new 'word.application'; Word.Visible = true
# WIN32OLE.const_load(Word, WORD)

def fine_print file
	file = file + (File.extname(file)=='' ? '.ppt' : '')
	case File.extname(file)
		when '.ppt'
			fine_print_ppt file
		when '.doc'
			fine_print_word file
		else
			raise 'Unknown File Type: ' + file
	end
end

def fine_print_ppt file
	# ppt_app = WIN32OLE.new 'powerpoint.application'
	presentation = Ppt.Presentations.open file
	# presentation = ppt_app.ActivePresentation

	print = presentation.PrintOptions
	print.activePrinter = 'FinePrint'
	print.rangeType = PPT::PpPrintAll
	print.numberOfCopies = 1
	print.collate = true
	print.outputType = PPT::PpPrintOutputSlides
	print.printHiddenSlides = PPT::PpPrintColor
	print.fitToPage = false
	print.handoutOrder = PPT::PpPrintHandoutHorizontalFirst
	
	presentation.printOut
	sleep 1
	presentation.close
end

def fine_print_word file
	doc = Word.documents.open file
	Word.activePrinter = 'FinePrint'
	doc.printOut
	sleep 1
	doc.close
end

at_exit {
	Ppt.quit
	Word.quit
}
