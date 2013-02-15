require 'digest/md5'

module SimpleXlsx
  class Hyperlinks

    attr_reader :io, :count

    def initialize io, sheet
      @io = io
      @sheet = sheet

      @ids = {}
    end

    def add_hyperlink x, y, target
      digest = Digest::MD5.digest(target)

      id = @ids[digest] ||
        (@ids[digest] = @sheet.add_relationship Relationships::TYPE_HYPERLINK, target, {:TargetMode=>:External})

      @io.write "<hyperlink ref=\"#{Sheet.column_index x}#{y+1}\" r:id=\"#{id.to_xs}\"/>"
    end

    def to_stream stream
      stream.write "<hyperlinks>"
      Stream.copy(@io, stream)
      stream.write "</hyperlinks>"
    end

  end
end

