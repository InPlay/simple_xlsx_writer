module SimpleXlsx
  class Styles
    class NumFmts < Base

      NUM0 = 1
      NUM0_00 = 2

      DATE = 15
      TIME = 22

      AT = 49

      GENERAL = 100
      BOOLEAN = 101
      YYYYMMDD = 102
      YYYYMMDDHHMMSS = 103

      def to_stream stream
        stream.puts "<numFmts count=\"#{length}\">"
        @content.each{|(c, id)|
          stream.puts "<numFmt formatCode=\"#{c[:format_code].to_s.to_xs}\" numFmtId=\"#{c[:id]}\"/>"
        }
        stream.puts "</numFmts>"
      end

      private

      def validate o
        super
        raise ArgumentError, "No number format id specified" unless o[:id] 
        raise ArgumentError, "Number format id should be a Fixnum" unless o[:id].is_a? Fixnum
        raise ArgumentError, "No number format code specified" unless o[:format_code]
      end

    end
  end
end
