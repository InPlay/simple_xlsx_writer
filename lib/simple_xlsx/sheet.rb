require 'bigdecimal'
require 'time'

module SimpleXlsx

class Sheet
  attr_reader :name
  attr_accessor :rid

  attr_reader :merged_cells, :relationships, :relationships_file

  def initialize document, name, stream, &block
    @document = document
    @stream =  stream
    @name = name
    @row_ndx = 1
    @stream.write <<-ends
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheetData>
ends
    begin
      if block_given?
        yield self
      end
      @stream.write "</sheetData>"

      if @merged_cells
        src = @merged_cells_file
        @stream.write "<mergeCells count='#{@merged_cells.count}'>"
        Stream.copy(src, @stream)
        @stream.write "</mergeCells>"
      end 

      if @hyperlinks
        src = @hyperlinks_file
        @stream.write "<hyperlinks>"
        Stream.copy(src, @stream)
        @stream.write "</hyperlinks>"
      end

      @stream.write "</worksheet>"
    ensure
      @merged_cells_file.unlink if @merged_cells_file
      @hyperlinks_file.unlink if @hyperlinks_file
      @relationships_file.unlink if @relationships_file 
    end
  end

  def merge_cells x1, y1, x2, y2
    @merged_cells ||= begin
      @merged_cells_file = Tempfile.new('xlsx-merged-cells')
      MergedCells.new @merged_cells_file
    end
    @merged_cells.merge_cells x1, y1, x2, y2
  end

  def add_relationship type, target
    @relationships ||= begin
      @relationships_file = Tempfile.new('xlsx-sheet-rels')
      Relationships.new @relationships_file
    end
    @relationships.add_relationship type, target
  end

  def add_hyperlink x, y, target
    @hyperlinks ||= begin
      @hyperlinks_file = Tempfile.new('xlsx-hyperlinks')
      Hyperlinks.new @hyperlinks_file, self
    end
    @hyperlinks.add_hyperlink x, y, target
  end

  def add_row arry
    row = ["<row r=\"#{@row_ndx}\">"]
    arry.each_with_index do |value, col_ndx|
      kind, ccontent, cstyle = format_field_and_type_and_style value
      row << "<c r=\"#{Sheet.column_index(col_ndx)}#{@row_ndx}\" t=\"#{kind.to_s}\" s=\"#{cstyle}\">#{ccontent}</c>"
    end
    row << "</row>"
    @row_ndx += 1
    @stream.write(row.join())
  end

  def format_field_and_type_and_style value
    if value.is_a?(String)
      if @document.has_shared_strings?
        [:s, "<v>#{(@document.shared_strings << value)}</v>", 5] 
      else
        [:inlineStr, "<is><t>#{value.to_xs}</t></is>", 5]
      end
    elsif value.is_a?(BigDecimal)
      [:n, "<v>#{value.to_s('f')}</v>", 4]
    elsif value.is_a?(Float)
      [:n, "<v>#{value.to_s}</v>", 4]
    elsif value.is_a?(Numeric)
      [:n, "<v>#{value.to_s}</v>", 3]
    elsif value.is_a?(Date)
      [:n, "<v>#{Sheet.days_since_jan_1_1900(value)}</v>", 2]
    elsif value.is_a?(Time)
      [:n, "<v>#{Sheet.fractional_days_since_jan_1_1900(value)}</v>", 1]
    elsif value.is_a?(TrueClass) || value.is_a?(FalseClass)
      [:b, "<v>#{value ? '1' : '0'}</v>", 6]
    else
      [:inlineStr, "<is><t>#{value.to_s.to_xs}</t></is>", 5]
    end
  end

  def self.days_since_jan_1_1900 date
    @@jan_1_1904 ||= Date.parse("1904 Jan 1")
    (date - @@jan_1_1904).to_i + 1462 # http://support.microsoft.com/kb/180162
  end

  def self.fractional_days_since_jan_1_1900 value
    @@jan_1_1904_midnight ||= ::Time.utc(1904, 1, 1)
    ((value - @@jan_1_1904_midnight) / 86400.0) + #24*60*60
      1462 # http://support.microsoft.com/kb/180162
  end

  def self.abc
    @@abc ||= ('A'..'Z').to_a
  end

  def self.column_index n
    result = []
    while n >= 26 do
      result << abc[n % 26]
      n /= 26
    end
    result << abc[result.empty? ? n : n - 1]
    result.reverse.join
  end

end
end
