require 'digest/md5'

module SimpleXlsx
  class SharedStrings

    attr_reader :io

    def initialize io
      @content = {}

      @count = 0
      @index = 0

      @io = io
    end

    def << str
      s = str.to_s
      @count += 1
      digest = Digest::MD5.digest(s)

      @content[digest] || begin
        result = @index

        @content[digest] = @index
        @io.puts "<si><t>#{s.to_xs}</t></si>"

        @index = @index + 1
        result
      end
    end

    def count
      @count
    end

    def unique_count
      @index
    end

    def to_stream stream
      stream.puts <<-ends
<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="#{count}" uniqueCount="#{unique_count}">
ends
      Stream.copy(@io, stream)
      stream.puts '</sst>' 
    end

  end
end

