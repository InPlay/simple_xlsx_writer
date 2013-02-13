module SimpleXlsx
  class Relationships

    attr_reader :io

    TYPE_STYLES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    TYPE_WORKSHEET = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
    TYPE_SHARED_STRINGS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
    TYPE_HYPERLINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"

    TARGET_STYLES = "styles.xml"
    TARGET_SHARED_STRINGS = "sharedStrings.xml"

    def initialize io
      @io = io
      @count = 0
    end

    def add_relationship type, target
      id = "rId#{@count += 1}"
      @io.puts "<Relationship Id=\"#{id.to_xs}\" Type=\"#{type.to_xs}\" Target=\"#{target.to_xs}\"/>"
      id
    end

  end
end


