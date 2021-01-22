import org.apache.poi.ss.usermodel.Row;

import java.util.function.Supplier;
import java.util.stream.IntStream;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

public class Util {
    public static Supplier<Stream<Row>> getRowStreamSupplier(Iterable<Row> rows) {
        return () -> getStream(rows);
    }

    public static <T> Stream<T> getStream(Iterable<T> iterable) {
        return StreamSupport.stream(iterable.spliterator(), false);
    }

    public static Supplier<Stream<Integer>> cellIteratorSupplier(int end) {
        return () -> numberStream(end);
    }

    public static Stream<Integer> numberStream(int end) {
        return IntStream.range(0, end).boxed();
    }

}
