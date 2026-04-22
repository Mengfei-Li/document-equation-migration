import io.transpect.calabash.extensions.Ole2XmlConverter;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;

public class Ole2XmlCli {
    public static void main(String[] args) throws IOException {
        if (args.length < 2) {
            System.err.println("Usage: Ole2XmlCli <input-bin-or-wmf> <output-xml>");
            System.exit(1);
        }

        Path input = Path.of(args[0]).toAbsolutePath().normalize();
        Path output = Path.of(args[1]).toAbsolutePath().normalize();

        String rubySafeInput = input.toString().replace("\\", "/");

        Ole2XmlConverter converter = new Ole2XmlConverter();
        String xml = converter.convertFormula(rubySafeInput);

        Files.createDirectories(output.getParent());
        Files.writeString(output, xml, StandardCharsets.UTF_8);
    }
}
