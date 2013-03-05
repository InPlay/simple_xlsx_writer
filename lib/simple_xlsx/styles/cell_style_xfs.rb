module SimpleXlsx
  class Styles
    class CellStyleXfs < Base

      def to_stream stream
        stream.puts "<cellStyleXfs count=\"#{length}\">"
        content.each{|(c,id)| Xf.write stream, c}
        stream.puts "</cellStyleXfs>"
      end

    end
  end
end

