module SimpleXlsx
  class ContentTypes

    attr_reader :io

    CONTENT_TYPE_RELATIONSHIPS = "application/vnd.openxmlformats-package.relationships+xml"
    CONTENT_TYPE_CORE_PROPERTIES = "application/vnd.openxmlformats-package.core-properties+xml"
    CONTENT_TYPE_EXT_PROPERTIES = "application/vnd.openxmlformats-officedocument.extended-properties+xml"
    CONTENT_TYPE_WORKBOOK = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
    CONTENT_TYPE_SHEET = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
    CONTENT_TYPE_SHARED_STRINGS = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
    CONTENT_TYPE_STYLES = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
    CONTENT_TYPE_THEME = "application/vnd.openxmlformats-officedocument.theme+xml"

    def initialize io
      @io = io
      @io.puts <<-eos
<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
eos
    end

    def add_content_type part_name, content_type
      @io.puts "<Override PartName=\"#{part_name.to_xs}\" ContentType=\"#{content_type.to_xs}\"/>"
    end

    def close
      @io.puts '</Types>'
    end

  end
end
