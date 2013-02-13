module SimpleXlsx
  module Stream
    class << self
      def copy src, dst
        src.flush
        src.rewind
        while buff = src.read(0x4000)
          dst.write(buff)
        end
      end
    end
  end
end

