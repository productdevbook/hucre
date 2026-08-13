// Validate an XML file against an XSD, reporting every error rather than
// stopping at the first.
//
// Java's built-in `javax.xml.validation` is the only XSD 1.0 validator
// this repository can reach: `xmllint` is not installed, and the npm
// options need a native build. Thirty lines of Java is a smaller ask
// than either, and `scripts/validate-ooxml.mjs` compiles it on demand —
// nothing is checked in but the source.
//
// Used by that script; not part of `pnpm test`, which must not need a JVM.

import javax.xml.XMLConstants;
import javax.xml.transform.stream.StreamSource;
import javax.xml.validation.Schema;
import javax.xml.validation.SchemaFactory;
import javax.xml.validation.Validator;
import org.xml.sax.ErrorHandler;
import org.xml.sax.SAXParseException;
import java.io.File;

/** Validate an XML file against an XSD. Prints every error, not just the first. */
public class XsdValidate {
  static int errors = 0;

  public static void main(String[] args) throws Exception {
    SchemaFactory factory = SchemaFactory.newInstance(XMLConstants.W3C_XML_SCHEMA_NS_URI);
    Schema schema = factory.newSchema(new File(args[0]));
    Validator validator = schema.newValidator();
    validator.setErrorHandler(new ErrorHandler() {
      public void warning(SAXParseException e) {}
      public void error(SAXParseException e) { report("error", e); }
      public void fatalError(SAXParseException e) { report("fatal", e); }
      void report(String kind, SAXParseException e) {
        if (errors++ < 20) System.out.println(kind + " " + e.getLineNumber() + ":" + e.getColumnNumber() + ": " + e.getMessage());
      }
    });
    try { validator.validate(new StreamSource(new File(args[1]))); }
    catch (SAXParseException e) { /* already reported */ }
    System.out.println(errors == 0 ? "VALID" : errors + " error(s)");
    System.exit(errors == 0 ? 0 : 1);
  }
}
