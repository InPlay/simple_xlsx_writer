require 'simple_xlsx/styles/base'
require 'simple_xlsx/styles/xf'
require 'simple_xlsx/styles/num_fmts'
require 'simple_xlsx/styles/fonts'
require 'simple_xlsx/styles/fills'
require 'simple_xlsx/styles/borders'
require 'simple_xlsx/styles/cell_style_xfs'
require 'simple_xlsx/styles/cell_xfs'
require 'simple_xlsx/styles/cell_styles'

module SimpleXlsx
  class Styles

    DEFAULT_FONT = 'Arial'
    DEFAULT_FONT_FAMILY = 0
    DEFAULT_FONT_SIZE = 10

    def initialize
      @num_fmts = NumFmts.new
      @fonts = Fonts.new
      @fills = Fills.new
      @borders = Borders.new
      @cell_style_xfs = CellStyleXfs.new
      @cell_xfs = CellXfs.new
      @cell_styles = CellStyles.new

      add_minimal
    end

    def add_style style
      num_fmt = find_num_fmt style
      font = find_font style
      fill = find_fill style
      border = find_border style

      alignment = style[:alignment]

      to_find = {:font_id=>font, :fill_id=>fill, :border_id=>border, :num_fmt_id=>num_fmt,
                      :apply_alignment=> !!alignment,
                      :apply_border=>true,
                      :apply_font=>true,
                      :apply_protection=>false, :xf_id=>0 }
      to_find.merge! :alignment=>alignment if alignment
      (@cell_xfs << to_find)
    end
    alias :<< :add_style

    PARTS = [:num_fmts, :fonts, :fills, :borders, :cell_style_xfs,
       :cell_xfs, :cell_styles]

    def to_stream stream
      stream.puts <<-eos
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
eos
      PARTS.each{|sym| (instance_variable_get "@#{sym}").to_stream stream}
      
      stream.puts '</styleSheet>'
    end

    private

    def find_num_fmt style
      n = style[:num_fmt]
      return (n || 0) if !n || n.is_a?(Integer)
      @num_fmts << n
    end

    def find_font style
      f = style[:font]
      return 0 if !f
      @fonts << f
    end

    def find_fill style
      f = style[:fill]
      return 0 if !f
      @fills << f
    end

    def find_border style
      b = style[:border]
      return 0 if !b
      @borders << b
    end

    def add_minimal
      @num_fmts << {:id => NumFmts::GENERAL, :format_code => "GENERAL"}
      @num_fmts << {:id => NumFmts::BOOLEAN, :format_code => "\"TRUE\";\"TRUE\";\"FALSE\""}
      @num_fmts << {:id => NumFmts::YYYYMMDD, :format_code => "yyyy/mm/dd"}
      @num_fmts << {:id => NumFmts::YYYYMMDDHHMMSS, :format_code=> "yyyy/mm/dd hh:mm:ss"}

      @fonts << {:name => DEFAULT_FONT, :size => DEFAULT_FONT_SIZE, :family => DEFAULT_FONT_FAMILY}

      @fills << {:pattern_fill => :none}
      @fills << {:pattern_fill => :gray125}

      @borders << {}
      @borders << Borders::BLACK

      @cell_xfs << {:font_id=>0, :fill_id=>0, :border_id=>0, :num_fmt_id=>NumFmts::NUM0,
                      :apply_border=>true,
                      :apply_font=>true,
                      :apply_protection=>false, :xf_id=>0 }

      @cell_style_xfs << {:border_id=>0, :num_fmt_id=>0, 
                          :font_id=>0, :fill_id=>0, :xf_id=>0}

      @cell_styles << {:name => 'Normal', :builtin_id =>0, :custom_builtin=>false, :xf_id=>0}
    end

  end
end

