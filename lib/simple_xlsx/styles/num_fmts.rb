module SimpleXlsx
  class Styles
    class NumFmts < Base

      NUM0 = 1
      NUM0_00 = 2

      CURRENCY = 5 # $#,##0_);($#,##0)

      DATE = 15
      TIME = 22

      AT = 49

      GENERAL = 100
      BOOLEAN = 101
      YYYYMMDD = 102
      YYYYMMDDHHMMSS = 103

      def initialize
        super
        @taken_ids = {}
        @next_id = 1000
      end

      def << o
        if o.is_a? Integer
          o
        elsif o.is_a? String
          add_num_fmt :format_code=>o
        elsif o.is_a? Hash
          add_num_fmt o
        else
          raise "Invalid number format, expected String, Hash or Integer."
        end
      end

      def to_stream stream
        stream.puts "<numFmts count=\"#{length}\">"
        @content.each{|(c, id)|
          stream.puts "<numFmt formatCode=\"#{c[:format_code].to_s.to_xs}\" numFmtId=\"#{id}\"/>"
        }
        stream.puts "</numFmts>"
      end

      private

      def add_num_fmt o
        id = o[:id]
        without_id = o.reject{|(k,v)| k == :id}

        already_added_id = @content[without_id]
        raise "Number format id #{id} already taken" if @taken_ids[id] && id != already_added_id

        validate without_id

        id ||= already_added_id
        id ||= begin
          r = @next_id
          @next_id = @next_id + 1
          r
        end

        @content[without_id] ||= id
        @taken_ids[id] = true

        id
      end 

      def validate o
        super
        raise ArgumentError, "No number format code specified" unless o[:format_code]
      end

    end
  end
end
