module SimpleXlsx
  class MergedCells

    attr_reader :io, :count

    def initialize io
      @count = 0
      @io = io
    end

    def merge_cells x1, y1, x2, y2
      xx, yy = [[x1, x2], [y1,y2]]
      minx, maxx = [xx.min, xx.max]
      miny, maxy = [yy.min+1, yy.max+1]
      @io.write "<mergeCell ref='#{Sheet.column_index minx}#{miny}:#{Sheet.column_index maxx}#{maxy}'/>"
      @count += 1
    end

    def to_stream stream
      stream.write "<mergeCells count='#{count}'>"
      Stream.copy(@io, stream)
      stream.write "</mergeCells>"
    end

  end
end
