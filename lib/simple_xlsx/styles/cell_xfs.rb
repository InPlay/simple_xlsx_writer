module SimpleXlsx
  class Styles
    class CellXfs < Base

      def to_stream stream
        stream.puts "<cellXfs count=\"#{length}\">"
        content.each{|(c,id)| Xf.write stream, c}
        stream.puts "</cellXfs>"
      end

    end
  end
end


