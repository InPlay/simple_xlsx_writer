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
      @count += 1
      digest = Digest::MD5.digest(str)

      @content[digest] || begin
        result = @index

        @content[digest] = @index
        @io.puts "<si><t>#{str.to_xs}</t></si>"

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

  end
end

