require 'zip'

module SimpleXlsx

class Serializer
  def initialize to
    @to = to
    Zip::File.open(to, Zip::File::CREATE) do |zip|

      @zip = zip

      @zip.mkdir "xl/_rels"
      @zip.get_output_stream "xl/_rels/workbook.xml.rels" do |relationships_file|
      @zip.get_output_stream "[Content_Types].xml" do |content_types_file|
        @content_types = ContentTypes.new content_types_file
        @relationships = Relationships.new relationships_file
        @styles = Styles.new

        add_content_types

        add_doc_props
        add_worksheets_directory
        add_relationship_part
        @doc = Document.new self, @content_types, @relationships, @styles
        yield @doc

        add_styles
        add_shared_strings if @doc.has_shared_strings?

        add_workbook_part
        @doc.close
      end
      end
    end
  end

  def add_workbook_part
    @zip.get_output_stream "xl/workbook.xml" do |f|
      f.puts <<-ends
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<workbookPr date1904="0" />
<sheets>
ends
      @doc.sheets.each_with_index do |sheet, ndx|
        # Excel fails to read sheet names longer than 31 character
        f.puts "<sheet name=\"#{sheet.name.to_s[0..30].to_xs}\" sheetId=\"#{ndx + 1}\" r:id=\"#{sheet.rid}\"/>"
      end
      f.puts "</sheets></workbook>"
    end
  end

  def add_worksheets_directory
    @zip.mkdir "xl"
    @zip.mkdir "xl/worksheets"
  end

  def open_stream_for_sheet ndx
    @zip.get_output_stream "xl/worksheets/sheet#{ndx + 1}.xml" do |f|
      yield f
    end
  end

  def open_stream_for_sheet_rels ndx
    if !@sheet_rels_created
      @zip.mkdir "xl/worksheets/_rels/"
      @sheet_rels_created = true
    end

    @zip.get_output_stream "xl/worksheets/_rels/sheet#{ndx + 1}.xml.rels"
  end

  def add_content_types
    @content_types.add_content_type '/_rels/.rels', ContentTypes::CONTENT_TYPE_RELATIONSHIPS
    @content_types.add_content_type '/docProps/core.xml', ContentTypes::CONTENT_TYPE_CORE_PROPERTIES
    @content_types.add_content_type '/docProps/app.xml', ContentTypes::CONTENT_TYPE_EXT_PROPERTIES
    @content_types.add_content_type '/xl/workbook.xml', ContentTypes::CONTENT_TYPE_WORKBOOK
    @content_types.add_content_type '/xl/_rels/workbook.xml.rels', ContentTypes::CONTENT_TYPE_RELATIONSHIPS
    @content_types.add_content_type '/xl/styles.xml', ContentTypes::CONTENT_TYPE_STYLES
  end

  def add_relationship_part
    @zip.mkdir "_rels"
    @zip.get_output_stream "_rels/.rels" do |f|
      f.puts <<-ends
<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>
ends
    end
  end

  def add_doc_props
    @zip.mkdir "docProps"
    @zip.get_output_stream "docProps/core.xml" do |f|
      f.puts <<-ends
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
   <dcterms:created xsi:type="dcterms:W3CDTF">2010-07-20T14:30:58.00Z</dcterms:created>
   <cp:revision>0</cp:revision>
</cp:coreProperties>
ends
    end
    @zip.get_output_stream "docProps/app.xml" do |f|
      f.puts <<-ends
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <TotalTime>0</TotalTime>
</Properties>
ends
    end
  end

  def add_styles
    @zip.get_output_stream "xl/styles.xml" do |f|
      @styles.to_stream f
    end
  end

  def add_shared_strings
    @zip.get_output_stream "xl/sharedStrings.xml" do |f|
      @doc.shared_strings.to_stream f
    end
  end

end

end

