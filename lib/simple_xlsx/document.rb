module SimpleXlsx
  class Document
    attr_accessor :use_shared_strings

    def initialize(io)
      @sheets = []
      @io = io
      @use_shared_strings = false

      @shared_strings_file = Tempfile.new('xlsx-shared-strings')
      @shared_strings = SharedStrings.new @shared_strings_file
    end

    attr_reader :sheets

    def add_sheet name, &block
      @io.open_stream_for_sheet(@sheets.size) do |stream|
        @sheets << Sheet.new(self, name, stream, &block)
      end
    end

    alias :has_shared_strings? :use_shared_strings

    attr_reader :shared_strings, :shared_strings_file

    def close
      @shared_strings_file.unlink
    end

  end
end
