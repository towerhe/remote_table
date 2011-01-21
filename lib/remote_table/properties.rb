require 'uri'
class RemoteTable
  class Properties
    attr_reader :t
    def initialize(t)
      @t = t
    end
    
    def uri
      return @uri if @uri.is_a?(::URI)
      @uri = ::URI.parse t.url
      if @uri.host == 'spreadsheets.google.com'
        @uri.query = 'output=csv&' + @uri.query.sub(/\&?output=.*?(\&|\z)/, '\1')
      end
      @uri
    end
    
    # The headers specified by the user
    def headers
      t.options['headers']
    end
    
    # The sheet specified by the user as a number or a string
    #
    # Default: 0
    def sheet
      t.options['sheet'] || 0
    end
    
    # Whether to keep blank rows
    #
    # Default: false
    def keep_blank_rows
      t.options['keep_blank_rows'] || false
    end
    
    # Form data to send in with the download request
    def form_data
      t.options['form_data']
    end
    
    # How many rows to skip
    def skip
      t.options['skip'].to_i
    end
    
    # The encoding
    #
    # Default: "UTF-8"
    def encoding
      t.options['encoding'] || 'UTF-8'
    end
    
    # The delimiter
    #
    # Default: ","
    def delimiter
      t.options['delimiter'] || ','
    end
    
    # The XPath used to find rows
    def row_xpath
      t.options['row_xpath']
    end
    
    # The XPath used to find columns
    def column_xpath
      t.options['column_xpath']
    end
    
    # Get the format in the form of RemoteTable::Format::Excel, etc.
    #
    # Note: treats all spreadsheets.google.com URLs as Format::Delimited (i.e., CSV)
    #
    # Default: tries to guess from file extension
    def format
      return Format::Delimited if uri.host == 'spreadsheets.google.com'
      clue = if t.options['format']
        t.options['format'].to_s
      else
        ::File.extname(uri.path)
      end
      case clue.downcase
      when /xlsx/, /excelx/
        Format::Excelx
      when /xls/, /excel/
        Format::Excel
      when /csv/, /tsv/, /delimited/
        Format::Delimited
      when /ods/, /open_?office/
        Format::OpenOffice
      when /fixed_?width/
        Format::FixedWidth
      when /htm/
        Format::HTML
      else
        raise Format::Unknown
      end
    end
  end
end