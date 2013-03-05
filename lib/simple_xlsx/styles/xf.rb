module SimpleXlsx
  class Styles
    module Xf

      def self.write stream, c
        stream.puts "<xf "
        stream.puts "     applyAlignment=\"#{c[:apply_alignment].to_s.to_xs}\"" if c.has_key? :apply_alignment
        stream.puts "     applyBorder=\"#{c[:apply_border].to_s.to_xs}\"" if c.has_key? :apply_border
        stream.puts "     applyFont=\"#{c[:apply_font].to_s.to_xs}\"" if c.has_key? :apply_font
        stream.puts "     applyProtection=\"#{c[:apply_protection].to_s.to_xs}\"" if c.has_key? :apply_protection
        stream.puts "     borderId=\"#{c[:border_id]}\"" 
        stream.puts "     fillId=\"#{c[:fill_id]}\"" 
        stream.puts "     fontId=\"#{c[:font_id]}\"" 
        stream.puts "     numFmtId=\"#{c[:num_fmt_id]}\""
        stream.puts "     xfId=\"#{c[:xf_id]}\"" if c[:xf_id]
        stream.puts "     />"
      end

    end
  end
end



