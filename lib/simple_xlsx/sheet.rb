require 'bigdecimal'
require 'time'

module SimpleXlsx

class Sheet
  include AttrsSerializer

  attr_reader :name
  attr_accessor :rid

  attr_reader :merged_cells

  def initialize opts, &block
    [:styles, :index, :io, :document, :stream, :name, :columns].each {|sym|
      self.instance_variable_set "@#{sym.to_s}".to_sym, opts.fetch(sym)
    }

    @document.content_types.add_content_type "/xl/worksheets/sheet#{@index + 1}.xml", ContentTypes::CONTENT_TYPE_SHEET
    self.rid = @document.relationships.add_relationship Relationships::TYPE_WORKSHEET, "worksheets/sheet#{@index + 1}.xml"
  
    @row_ndx = 1
    @stream.write <<-ends
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
ends

    unless @columns.empty?
      @stream.write "<cols>"
      @columns.each {|c|
        attrs = {
          :style => 0,
        }.merge({
          :min=>c[:min],
          :max=>c[:max],
          :collapsed=>!!c[:collapsed],
          :hidden=>!!c[:hidden],
          :bestFit=>!!c[:best_fit],
          :width=> (c[:width] || 9),
          :customWidth=>!!c[:custom_width]
        })
        @stream.write "<col #{serialize_attrs attrs}/>"
      }
      @stream.write "</cols>"
    end

    @stream.write "<sheetData>"

    begin
      if block_given?
        yield self
      end
      @stream.write "</sheetData>"

      @merged_cells.to_stream @stream if @merged_cells
      @hyperlinks.to_stream @stream if @hyperlinks

      @stream.write "</worksheet>"
    ensure
      @merged_cells_file.unlink if @merged_cells_file
      @hyperlinks_file.unlink if @hyperlinks_file
      @relationships.close if @relationships
      @relationships_file.close if @relationships_file
    end
  end

  def merge_cells x1, y1, x2, y2
    @merged_cells ||= begin
      @merged_cells_file = Tempfile.new('xlsx-merged-cells')
      MergedCells.new @merged_cells_file
    end
    @merged_cells.merge_cells x1, y1, x2, y2
  end

  def add_relationship type, target, attrs  ={}
    @relationships ||= begin
      @document.content_types.add_content_type "/xl/worksheets/_rels/sheet#{@index + 1}.xml.rels", ContentTypes::CONTENT_TYPE_RELATIONSHIPS
      @relationships_file = @io.open_stream_for_sheet_rels @index
      Relationships.new @relationships_file
    end
    @relationships.add_relationship type, target, attrs
  end

  def add_hyperlink x, y, target
    @hyperlinks ||= begin
      @hyperlinks_file = Tempfile.new('xlsx-hyperlinks')
      Hyperlinks.new @hyperlinks_file, self
    end
    @hyperlinks.add_hyperlink x, y, target
  end

  def add_row arry, defaults = {}
    row = [defaults[:height] ? "<row r=\"#{@row_ndx}\" ht=\"#{defaults[:height]}\" customHeight=\"1\">" : "<row r=\"#{@row_ndx}\">"]
    arry.each_with_index do |v, col_ndx|
      value = Sheet.deep_merge(defaults, (Sheet.value_to_hash v))

      kind, ccontent, cstyle_base = format_field_and_type_and_style value[:value]

      final_value = (Sheet.deep_merge({:style=>cstyle_base}, value))
      cstyle = (@styles << final_value[:style])

      link = final_value[:link]
      add_hyperlink col_ndx, @row_ndx-1, link if link

      row << "<c r=\"#{Sheet.column_index(col_ndx)}#{@row_ndx}\" t=\"#{kind.to_s}\" s=\"#{cstyle}\">#{ccontent}</c>"
    end
    row << "</row>"
    @row_ndx += 1
    @stream.write(row.join())
  end

  def format_field_and_type_and_style value
    if value.is_a?(String) || value.is_a?(Array)
      if @document.has_shared_strings?
        [:s, "<v>#{(@document.shared_strings << Sheet.serialize_rich_string(value))}</v>", {:num_fmt => Styles::NumFmts::AT}] 
      else
        [:inlineStr, "<is>#{Sheet.serialize_rich_string(value)}</is>", {:num_fmt => Styles::NumFmts::AT}]
      end
    elsif value.is_a?(BigDecimal)
      [:n, "<v>#{value.to_s('f')}</v>", {:num_fmt => Styles::NumFmts::NUM0_00}]
    elsif value.is_a?(Float)
      [:n, "<v>#{value.to_s}</v>", {:num_fmt => Styles::NumFmts::NUM0_00}]
    elsif value.is_a?(Numeric)
      [:n, "<v>#{value.to_s}</v>", {:num_fmt => Styles::NumFmts::NUM0}]
    elsif value.is_a?(Date)
      [:n, "<v>#{Sheet.days_since_jan_1_1900(value)}</v>", {:num_fmt => Styles::NumFmts::DATE}]
    elsif value.is_a?(Time)
      [:n, "<v>#{Sheet.fractional_days_since_jan_1_1900(value)}</v>", {:num_fmt => Styles::NumFmts::TIME}]
    elsif value.is_a?(TrueClass) || value.is_a?(FalseClass)
      [:b, "<v>#{value ? '1' : '0'}</v>", {:num_fmt => Styles::NumFmts::BOOLEAN}]
    else
      [:inlineStr, "<is><t>#{value.to_s.to_xs}</t></is>", {:num_fmt => Styles::NumFmts::AT}]
    end
  end

  def self.serialize_rich_string value
    if value.is_a?(String)
      "<t>#{value.to_xs}</t>"
    else
      value.map{|s| 
        "<r>#{Sheet.inline_style s}<t xml:space=\"preserve\">#{self.inline_value s}</t></r>"
      }.join
    end  
  end

  def self.inline_style s
    if s.is_a?(String)
      ''
    else
      bold = !!s[:bold]
      italic = !!s[:italic]
      underline = !!( s[:underline] || s[:underlined] )

      font = s[:font] || Styles::DEFAULT_FONT
      font_family = s[:font_family] || Styles::DEFAULT_FONT_FAMILY
      font_size = s[:font_size] || Styles::DEFAULT_FONT_SIZE

      r = ["<rPr>"]
      r << "<b/>" if bold
      r << "<i/>" if italic
      r << "<u/>" if underline
      r << "<rFont val=\"#{font.to_xs}\"/><family val=\"#{font_family.to_s.to_xs}\"/><sz val=\"#{font_size.to_s.to_xs}\"/>"
      r << "<color rgb=\"#{SimpleXlsx::Styles::Base.format_color s[:color]}\"/>" if s[:color] 
      r << "</rPr>"

      r.join
    end
  end

  def self.inline_value s
    if s.is_a?(String)
      s.to_xs
    elsif s.is_a?(Hash)
      (s[:value] || '').to_s.to_xs
    else
      s.to_s.to_xs
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

  def self.value_to_hash v
    return v if v && v.is_a?(Hash)
    {:value=>v}
  end

  def self.deep_merge o1, o2
    o1.merge(o2) do |key, oldval, newval|
      oldval = oldval.to_hash if oldval.respond_to?(:to_hash)
      newval = newval.to_hash if newval.respond_to?(:to_hash)
      oldval.class.to_s == 'Hash' && newval.class.to_s == 'Hash' ? deep_merge(oldval, newval) : newval
    end
  end

end
end
