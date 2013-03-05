module SimpleXlsx
  class Styles
    class Borders < Base

      BORDER_ELEMENTS = [:left, :right, :top, :bottom]

      DEFAULT = {
        :diagonal_down => false,
        :diagonal_up => false,
        :outline => false,
      }

      BLACK = DEFAULT.merge(
        BORDER_ELEMENTS.inject({}) do |r, name|
          r[name] = {:style=>:thin, :color=>'ff000000'}
          r
        end
      )

      def to_stream stream
        stream.puts "<borders count=\"#{length}\">"
        content.each{|(c,id)|
          stream.puts "<border diagonal_up=\"#{c[:diagonal_up].to_s.to_xs}\""
          stream.puts "  diagonal_down=\"#{c[:diagonal_down].to_s.to_xs}\""
          stream.puts "  outline=\"#{c[:outline].to_s.to_xs}\">"

          [:left, :right, :top, :bottom].map{|e| c[e] ? [e,c[e]] : nil}.compact.each{|(e,c)|
            stream.puts "<#{e.to_s.to_xs} style=\"#{(c[:style].to_s || 'none').to_xs}\">"
            stream.puts "  <color rgb=\"#{Base.format_color c[:color]}\"/>" if c[:color]
            stream.puts "</#{e.to_s.to_xs}>"
          }
          
          stream.puts "</border>"
        }
        stream.puts "</borders>"
      end

      private

      def apply_defaults hash
        DEFAULT.merge hash
      end    

      def validate o
        super
        [:left, :right, :top, :bottom].each {|e|
          if o[e]
            Base.validate_color o[e][:color] if o[e][:color]
          end
        }
      end

    end
  end
end
