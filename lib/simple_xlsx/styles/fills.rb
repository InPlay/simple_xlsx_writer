module SimpleXlsx
  class Styles
    class Fills < Base
    

      VALID_FILLS=[:none, :solid, :mediumGray, :darkGray,
        :lightGray, :darkHorizontal, :darkVertical,
        :darkDown, :darkUp, :darkGrid, :darkTrellis,
        :lightHorizontal, :lightVertical, :lightDown,
        :lightUp, :lightGrid, :lightTrellis, :gray125, :gray0625]

      def to_stream stream
        stream.puts "<fills count=\"#{length}\">"
        content.each{|(c,id)|
          if c[:pattern_fill]
            stream.puts "<fill>"
            stream.puts "  <patternFill patternType=\"#{c[:pattern_fill]}\">"
            stream.puts "    <fgColor rgb=\"#{Base.format_color c[:fg_color]}\"/>" if c[:fg_color]
            stream.puts "    <bgColor rgb=\"#{Base.format_color(c[:bg_color] || 'ff000000')}\"/>" if c[:fg_color]
            stream.puts "  </patternFill>"
            stream.puts "</fill>"
          end
        }
        stream.puts "</fills>"
      end

      private

      def validate o
        super
        if o[:pattern_fill]
          raise  ArgumentError, 
                "Invalid pattern_fill value, must be one of #{VALID_FILLS.map(&:to_s).join ', '}" unless VALID_FILLS.include? o[:pattern_fill]
        end
        Base.validate_color o[:fg_color] if o[:fg_color]
        Base.validate_color o[:bg_color] if o[:bg_color]
      end

    end
  end
end

