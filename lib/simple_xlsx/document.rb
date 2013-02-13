module SimpleXlsx
  class Document
    attr_accessor :use_shared_strings

    attr_reader :relationships, :relationships_file

    def initialize(io)
      @sheets = []
      @io = io
      @use_shared_strings = false

      @relationships_file = Tempfile.new('xlsx-shared-strings')
      @relationships = Relationships.new @relationships_file

      @relationships.add_relationship Relationships::TYPE_STYLES, Relationships::TARGET_STYLES

      @shared_strings_file = Tempfile.new('xlsx-rels')
      @shared_strings = SharedStrings.new @shared_strings_file
    end

    attr_reader :sheets

    def add_sheet name, &block
      idx = @sheets.size
      @io.open_stream_for_sheet(idx) do |stream|
        rid = @relationships.add_relationship Relationships::TYPE_WORKSHEET, 
          "worksheets/sheet#{idx+1}.xml"
        sheet = Sheet.new(self, name, stream, &block)
        sheet.rid = rid

        if sheet.relationships
          @io.open_stream_for_sheet_rels(idx) do |f|
            src = sheet.relationships_file
            Serializer.write_relationships src, f
          end
        end

        @sheets << sheet
      end
    end

    alias :has_shared_strings? :use_shared_strings

    attr_reader :shared_strings, :shared_strings_file

    def close
      @shared_strings_file.unlink
      @relationships_file.unlink
    end

  end
end
