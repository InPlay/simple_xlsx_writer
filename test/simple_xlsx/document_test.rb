require File.join(File.dirname(__FILE__), '..', 'test_helper')

module SimpleXlsx

class DocumentTest < Test::Unit::TestCase

  def open_stream_for_sheet sheets_size
    assert_equal sheets_size, @doc.sheets.size
    yield self
  end

  def write arg
  end

  def test_add_sheet
    @content_types = ContentTypes.new(StringIO.new "")
    @relationships = Relationships.new(StringIO.new "")
    @styles = Styles.new()

    @doc = Document.new self, @content_types, @relationships, @styles
    assert_equal [], @doc.sheets
    @doc.add_sheet "new sheet"
    assert_equal 1, @doc.sheets.size
    assert_equal 'new sheet', @doc.sheets.first.name
  end

end
end
