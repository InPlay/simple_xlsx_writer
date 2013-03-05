module SimpleXlsx
  class Styles
    module Xf

      def self.write stream, c
        stream.write "<xf"
        stream.write " applyAlignment=\"#{c[:apply_alignment].to_s.to_xs}\"" if c.has_key? :apply_alignment
        stream.write " applyBorder=\"#{c[:apply_border].to_s.to_xs}\"" if c.has_key? :apply_border
        stream.write " applyFont=\"#{c[:apply_font].to_s.to_xs}\"" if c.has_key? :apply_font
        stream.write " applyProtection=\"#{c[:apply_protection].to_s.to_xs}\"" if c.has_key? :apply_protection
        stream.write " borderId=\"#{c[:border_id]}\"" 
        stream.write " fillId=\"#{c[:fill_id]}\"" 
        stream.write " fontId=\"#{c[:font_id]}\"" 
        stream.write " numFmtId=\"#{c[:num_fmt_id]}\""
        stream.write " xfId=\"#{c[:xf_id]}\"" if c[:xf_id]
        stream.puts ">"

        stream.write "<alignment wrapText=\"true\"/>" if c[:alignment] && c[:alignment][:wrap_text]

        stream.puts "</xf>"
      end

    end
  end
end



