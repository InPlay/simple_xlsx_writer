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
        (@ids[digest] = @sheet.add_relationship Relationships::TYPE_HYPERLINK, target)

      @io.write "<hyperlink ref='#{Sheet.column_index x}#{y+1}' r:id='#{id.to_xs}'/>"
    end

  end
end

