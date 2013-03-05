module SimpleXlsx
  class Styles
    class Fonts < Base

      DEFAULT = {:name=>'Arial', :family=>0, :size=>10}

      def to_stream stream
        stream.puts "<fonts count=\"#{length}\">"
        content.each{|(c,id)|
          stream.puts "<font>"
          stream.puts "  <name val=\"#{c[:name].to_xs}\"/>"
          stream.puts "  <family val=\"#{c[:family].to_s.to_xs}\"/>"
          stream.puts "  <sz val=\"#{c[:size].to_s.to_xs}\"/>"
          stream.puts "  <color rgb=\"#{Base.format_color c[:color]}\"/>" if c[:color] 
          stream.puts "  <b/>" if c[:bold]
          stream.puts "  <i/>" if c[:italic]
          stream.puts "  <u/>" if c[:underline] || c[:underlined]
          stream.puts "</font>"
        }
        stream.puts "</fonts>"
      end

      private

      def apply_defaults hash
        DEFAULT.merge hash
      end

    end
  end
end
