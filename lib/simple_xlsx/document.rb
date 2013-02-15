module SimpleXlsx
  class Document
    attr_accessor :use_shared_strings, :content_types, :relationships

    def initialize(io, content_types, relationships)
      @sheets = []
      @io = io
      @use_shared_strings = false

      @content_types = content_types

      @relationships = relationships
      @relationships.add_relationship Relationships::TYPE_STYLES, Relationships::TARGET_STYLES
    end

    attr_reader :sheets

    def add_sheet name, &block
      idx = @sheets.size
      @io.open_stream_for_sheet(idx) do |stream|
        @sheets << Sheet.new(:io=>@io, 
                              :document=>self, 
                              :name=>name, 
                              :stream=>stream,
                              :index=>idx,
                              &block)
      end
    end

    def has_shared_strings?
      !!@use_shared_strings
    end

    def shared_strings
      @shared_strings ||= begin
        @content_types.add_content_type "/xl/sharedStrings.xml", ContentTypes::CONTENT_TYPE_SHARED_STRINGS
        @relationships.add_relationship  Relationships::TYPE_SHARED_STRINGS, "sharedStrings.xml" 
        @shared_strings_file = Tempfile.new('xlsx-shared-strings')
        SharedStrings.new @shared_strings_file
      end
    end

    def close
      @shared_strings_file.unlink if @shared_strings_file
      @relationships.close
      @content_types.close
    end

  end
end
