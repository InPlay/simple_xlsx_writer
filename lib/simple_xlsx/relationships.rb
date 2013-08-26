module SimpleXlsx
  class Relationships

    attr_reader :io

    TYPE_STYLES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    TYPE_WORKSHEET = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
    TYPE_SHARED_STRINGS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
    TYPE_HYPERLINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
    TYPE_THEME = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"

    TARGET_STYLES = "styles.xml"
    TARGET_SHARED_STRINGS = "sharedStrings.xml"

    def initialize io
      @io = io
      @count = 0

      @io.puts <<-ends
<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
ends
    end

    def add_relationship type, target, attrs = {}
      id = "rId#{@count += 1}"
      @io.write "<Relationship Id=\"#{id.to_xs}\" Type=\"#{type.to_xs}\" Target=\"#{ExcelCompatibility::truncate_uri(target).to_xs}\""
      attrs.each{|k,v| @io.write " #{k.to_s}=\"#{v.to_s.to_xs}\""}
      @io.write "/>"
      id
    end

    def close
      @io.puts "</Relationships>"
    end

  end
end
