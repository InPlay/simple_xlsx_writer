module SimpleXlsx
  module CopyStream

    def copy_stream src, dst
      src.flush
      src.rewind
      while buff = src.read(0x4000)
        dst.write(buff)
      end
    end

  end
end

