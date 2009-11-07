#!/usr/bin/env ruby -wKU

require 'getoptlong'
require 'pdf/reader'
require 'zip/zip'
require 'zip/zipfilesystem'

if RUBY_PLATFORM =~ /mingw|mswin/
  require 'win32ole'
  require 'Win32API'
  class Wcol
    gsh = Win32API.new("kernel32", "GetStdHandle", ['L'], 'L') 
    @textAttr = Win32API.new("kernel32","SetConsoleTextAttribute", ['L','N'], 'I')
    @h = gsh.call(-11)

    def self.color(col)
      @textAttr.call(@h,col)
    end
  end
end

class PageTextReceiver
   attr_accessor :content

   def initialize
     @content = []
   end

   # Called when page parsing starts
   def begin_page(arg = nil)
     @content << ""
   end

   # record text that is drawn on the page
   def show_text(string, *params)
     @content.last << string.strip
   end

   # there's a few text callbacks, so make sure we process them all
   alias :super_show_text :show_text
   alias :move_to_next_line_and_show_text :show_text
   alias :set_spacing_next_line_show_text :show_text

   # this final text callback takes slightly different arguments
   def show_text_with_positioning(*params)
     params = params.first
     params.each { |str| show_text(str) if str.kind_of?(String)}
   end
end

class Docs
  def initialize(dir = ".", query = nil, debug = false)
    @dir = dir
    @query = query
    @debug = debug
    @documents = Queue.new
    @emails = []
    @results = []
    @lock = Mutex.new
  end
  
  def find
   files = Dir["#{@dir}/**/*.*"]
   files.select {|x| /.pdf$|.doc$|.docx$|.xlsx$|.pptx$|.odt$|.odp$|.ods$|.odb$|.txt$|.rtf$|.ans$|.csv$|.xml|.json$|.html$/i}.each { |f| push(f) } 
  end
  
  def search!
    while @documents.size >=1
      @threads << Thread.new { detect_type }
      @threads.each {|t| t.join } if @threads != nil
    end
  end
  
  def search
    detect_type while @documents.size >=1
  end
  
  private
  ### HELPERS ###  
  def push(doc)
    @@document.push(doc)
  end
  
  def D(msg)
    puts msg if @debug
  end
  
  ### PARSERS ###
  def detect_type
    document = @documents.pop
    puts "\tParsing #{document}\n"
    case document
    when /.pdf/
      pdf(name)
    when /.doc/
      doc(name)
    when /.txt|.rtf|.ans|.csv|.html|.json/
      plain(name)
    when /.docx|.xlsx|.pptx|.odt|.odp|.ods|.odb/
      zxml(name)
    else
    end
  end
    
  end
  
  def pdf(name)
    begin
      receiver = PageTextReceiver.new
      pdf = PDF::Reader.file(name, receiver)
      search_emails(receiver.content.inspect)
    rescue PDF::Reader::UnsupportedFeatureError
      D "Error: Encrypted PDF - Unable to parse.\n"
    rescue PDF::Reader::MalformedPDFError
      D "Error: Malformed PDF - Unable to parse.\n"
    rescue
      D "Error: Unknown - Unable to parse.\n"
    end
  end

  def doc(name)
    if RUBY_PLATFORM =~ /mingw|mswin/
      begin
        word(name)
      rescue
        antiword(name)
      end
    elsif RUBY_PLATFORM =~ /linux|darwin/
      begin
        antiword(name)
      rescue
        D "Error: Unable to parse .doc"
      end
    else
      D "Error: Platform not supported."
    end
  end

  def word(name)
    word = WIN32OLE.new('word.application')
    word.documents.open(name)
    word.selection.wholestory
    search_emails(word.selection.text.chomp)
    word.activedocument.close( false )
    word.quit
  end

  def antiword(name)
    case RUBY_PLATFORM
    when /mingw|mswin/
      if File.exists?("C:\\antiword\\antiword.exe")
        search_emails(`C:\\antiword\\antiword.exe "#{name}" -f -s`) 
      end
    when /linux|darwin/
      if File.exists?("/usr/bin/antiword") or 
         File.exists?("/usr/local/bin/antiword") or 
         File.exists?("/opt/local/bin/antiword")
        search_emails(`antiword "#{name}" -f -s`) 
      end
    else
       # This G h e t t o but, for now it works on emails 
       # that do not contain Capital letters:)
       D "Debug: Using the Ghetto way."
       search_emails( File.open(name).readlines[0..19].to_s )
    end
  end

  def plain(data)
    search_emails(File.open(name).readlines.to_s)
  end

  def zxml(name)
    begin
      Zip::ZipFile.open(name) do |zip|
        text = z.entries.each { |e| zip.file.read(e.name) if e.name =~ /.xml$/}
        search_emails(text)
      end
    rescue
      D "Error: Unable to parse .#{name.scan(/\..[a-z]*$/)}\n"
    end
  end
  
  ### PARSE FOR EMAILS ###
  def search_emails(string,name)
    list = string.scan(/[a-z0-9!#$&'*+=?^_`{|}~-]+(?:\.[a-z0-9!#$&'*+=?^_`{|}~-]+)*_at_\
  (?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z](?:[a-z-]*[a-z])?|\
  [a-z0-9!#$&'*+=?^_`{|}~-]+(?:\.[a-z0-9!#$&'*+=?^_`{|}~-]+)*\sat\s(?:[a-z0-9](?:[a-z0-9-]\
  *[a-z0-9])?\.)+[a-z](?:[a-z-]*[a-z])?|[a-z0-9!#$&'*+=?^_`{|}~-]+\
  (?:\.[a-z0-9!#$&'*+=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z](?:[a-z-]*[a-z])?|\
  [a-z0-9!#$&'*+=?^_`{|}~-]+(?:\.[a-z0-9!#$&'*+=?^_`{|}~-]+)*\s@\s(?:[a-z0-9](?:[a-z0-9-]*\
  [a-z0-9])?\.)+[a-z](?:[a-z-]*[a-z])?|[a-z0-9!#$&'*+=?^_`{|}~-]+(?:\sdot\s[a-z0-9!#$&'*+=?^_`\
  {|}~-]+)*\sat\s(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\sdot\s)+[a-z](?:[a-z-]*[a-z])??/i)
    @lock.synchronize do
      print_(list)
      c_list = fix(list)
      @emails.concat(c_list).uniq!
      c_list.zip do |e| 
        @results << [e[0], name, e[0].match(/#{@query.gsub("@","").split('.')[0]}/) ? "T" : "F"]
      end
    end
  end
  
  ### FIX EMAILS ###
  def fix(list)
    list.each do |e|
      e.gsub!(" at ","@")
      e.gsub!("_at_","@")
      e.gsub!(" dot ",".")
      e.gsub!(/[+0-9]{0,3}[0-9()]{3,5}[-]{0,1}[0-9]{3,4}[-]{0,1}[0-9]{3,5}/,"")
    end
  end
  
  ### PRINTING FINDING METHODS ###
  def print_(list)
    list.each do |email|
      unless @emails.include?(email)
        case RUBY_PLATFORM
        when /mingw|mswin/
          print_windows
        when /linux|darwin/
          print_linux
        end
      end
    end
  end
  
  def print_linux(email)
    if email.match(/#{@query.gsub("@","").split('.')[0]}/)
      puts "\033[31m" + email + "\033\[0m"
    else
      puts "\033[32m" + email + "\033\[0m"
    end
  end
  
  def print_windows(email)
    if email.match(/#{@query.gsub("@","").split('.')[0]}/)
      Wcol::color(12)
      puts email
      Wcol::color(7)
    else
      Wcol::color(2)
      puts email
      Wcol::color(7)
    end
  end
  
  ### SAVING TO DISK ###
  def save(output)
    case output
    when /pdf/
      save_pdf
    when /csv/
      save_csv
    when /sqlite/
      save_sqlite
    end
  end
  
  def save_csv(name = "output.csv")
    out = File.new(name, "w")
    out << "EMAILS,DOCUMENTS, MATCHES QUERY\n"
    @results.each do |r|
      out << "#{r[0]},#{r[1]}\n,#{r[2]}"
    end
    
  end
  
  def save_pdf(name = "output.pdf")
    require 'prawn'
    require 'prawn/layout'
    Prawn::Document.generate(name) do  
      table @results, 
        :position => :center, 
        :headers => ["Email Address", "Document", "Matches query?"],
        :header_color => "0046f9",
        :row_colors => :pdf_writer, #["ffffff","ffff00"],
        :font_size => 10,
        :vertical_padding => 2,
        :horizontal_padding => 5
  end
  
  def save_sqlite
    require 'sqlite3'
    @db = SQLite3::Database.new(file)
    @db.execute("CREATE TABLE IF NOT EXISTS results (
      id integer primary key asc, 
      email text, 
      document text, 
      match char);")
      
    @results.each do |r| 
      @db.execute("INTERT INTO results (domain,email,score) VALUES (#{r[0]},#{r[1]},#{r[2]});")
    end
    @db.commit        
  end
end


opts = GetoptLong.new(
[ '--help', '-h', GetoptLong::NO_ARGUMENT ],
['--dir','-d', GetoptLong::REQUIRED_ARGUMENT ],
['--file','-f', GetoptLong::REQUIRED_ARGUMENT ],
['--pattern','-p', GetoptLong::REQUIRED_ARGUMENT ],
['--output','-o', GetoptLong::REQUIRED_ARGUMENT],
['--debug','-D', GetoptLong::NO_ARGUMENT ],
['--thread','-T', GetoptLong::NO_ARGUMENT ]
)

opts.each do |opt, arg|
  case opt
  when '--help':
    # BEGIN OF HELP
    puts "\nHELP for LSearchy\n---------------------\n
    --help, -h
    \tWell I guess you know what this is for (To obtain this Help).\n
    --dir, -d [directory_name]
    \t The root path to start the files.\n
    --query, -q
    \t The pattern to use to detect emails (i.e *client*).\n
    --output, -o
    \tThe output file name.
    Copyright 2009 - FreedomCoder\n"
    #END OF HELP
    exit(0)
  when '--dir':
    if File.exists?(arg)
      @dir = arg
    else
      puts "Directory not found"
    end  
  when '--query':
    @query = arg
  when '--output':
    @output = arg
    @type = @output.split(".")[1].upcase
    if  !(@type =~ /pdf|csv|sqlite/i)
      puts "unrecognized file type. bye!"
      exit(0)
    end
  when '--debug':
    @debug = true
  when '--thread':
    @thread = true
  else
    puts "Unknown command. Please try again"
    exit(0)
  end
end


if @dir 
  s = Docs.new(@dir,@query,@debug)
  @thread ? s.search_threaded : s.search
  s.save(@output)
end

puts "Bye! :)"
