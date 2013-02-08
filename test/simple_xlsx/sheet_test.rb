require File.join(File.dirname(__FILE__), '..', 'test_helper')
require "rexml/document"
require 'time'

module SimpleXlsx

class SheetTest < Test::Unit::TestCase

  def test_column_index
    assert_equal 'A', Sheet.column_index(0)
    assert_equal 'B', Sheet.column_index(1)
    assert_equal 'C', Sheet.column_index(2)
    assert_equal 'D', Sheet.column_index(3)
    assert_equal 'Y', Sheet.column_index(24)
    assert_equal 'Z', Sheet.column_index(25)
  end

  def test_column_index_two_digits
    assert_equal 'AA', Sheet.column_index(0+26)
    assert_equal 'AB', Sheet.column_index(1+26)
    assert_equal 'AC', Sheet.column_index(2+26)
    assert_equal 'AD', Sheet.column_index(3+26)
    assert_equal 'AZ', Sheet.column_index(25+26)
    assert_equal 'BA', Sheet.column_index(25+26+1)
    assert_equal 'BB', Sheet.column_index(25+26+2)
    assert_equal 'BC', Sheet.column_index(25+26+3)
  end

  def empty_sheet use_shared_strings = false 
    sheet_str = ""
    doc_str = ""
    io = StringIO.new sheet_str
    doc_io = StringIO.new doc_str

    doc = Document.new io
    doc.use_shared_strings = true if use_shared_strings
    Sheet.new(doc, 'name', io)
  end

  def test_format_field_for_strings
    v = empty_sheet.format_field_and_type_and_style "<escape this>"
    assert_equal [:inlineStr, "<is><t>&lt;escape this&gt;</t></is>", 5], v
  end

  def test_format_field_for_shared_strings
    sheet = empty_sheet true

    v = sheet.format_field_and_type_and_style "frequent string"
    assert_equal [:s, "<v>0</v>", 5], v

    v = sheet.format_field_and_type_and_style "rare"
    assert_equal [:s, "<v>1</v>", 5], v

    v = sheet.format_field_and_type_and_style "frequent string"
    assert_equal [:s, "<v>0</v>", 5], v
  end

  def test_format_field_for_numbers
    v = empty_sheet.format_field_and_type_and_style 3
    assert_equal [:n, "<v>3</v>", 3], v
    v = empty_sheet.format_field_and_type_and_style(BigDecimal.new("45"))
    assert_equal [:n, "<v>45.0</v>", 4], v
    v = empty_sheet.format_field_and_type_and_style(9.32)
    assert_equal [:n, "<v>9.32</v>", 4], v
  end

  def test_format_field_for_date
    v = empty_sheet.format_field_and_type_and_style(Date.parse('2010-Jul-24'))
    assert_equal [:n, "<v>#{38921+1462}</v>", 2], v
  end

  def test_format_field_for_datetime
    v = empty_sheet.format_field_and_type_and_style(Time.parse('2010-Jul-24 12:00 UTC'))
    assert_equal [:n, "<v>#{38921.5+1462}</v>", 1], v
  end


  def test_format_field_for_boolean
    v = empty_sheet.format_field_and_type_and_style(false)
    assert_equal [:b, "<v>0</v>", 6], v
    v = empty_sheet.format_field_and_type_and_style(true)
    assert_equal [:b, "<v>1</v>", 6], v
  end

  def test_add_row
    str = ""
    io = StringIO.new(str)
    doc_str = ""
    doc_io = StringIO.new doc_str
    sheet_doc = Document.new io

    Sheet.new(sheet_doc, 'name', io) do |sheet|
      sheet.add_row ['this is ', 'a new row']
    end
    doc = REXML::Document.new str
    assert_equal 'worksheet', doc.root.name
    sheetdata = doc.root.elements['sheetData']
    assert sheetdata
    row = sheetdata.elements['row']
    assert row
    assert_equal '1', row.attributes['r']
    assert_equal 2, row.elements.to_a.size
    attribute_keys = row.elements.to_a[0].attributes.keys
    assert attribute_keys.include?('r')
    assert attribute_keys.include?('s')
    assert attribute_keys.include?('t')
  end


end

end
