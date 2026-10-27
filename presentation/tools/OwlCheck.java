import java.io.File;
import org.semanticweb.owlapi.apibinding.OWLManager;
import org.semanticweb.owlapi.formats.*;
import org.semanticweb.owlapi.io.*;
import org.semanticweb.owlapi.model.*;
import org.semanticweb.owlapi.profiles.*;

/** Parse a Manchester Syntax file strictly, report OWL 2 DL profile violations, print axioms. */
public class OwlCheck {
  public static void main(String[] a) throws Exception {
    OWLOntologyManager m = OWLManager.createOWLOntologyManager();
    m.getOntologyConfigurator().setStrict(true);
    OWLOntology o = m.loadOntologyFromOntologyDocument(
        new FileDocumentSource(new File(a[0]), new ManchesterSyntaxDocumentFormat()));
    System.out.println("PARSED OK: " + o.getAxiomCount() + " axioms, " + o.getLogicalAxiomCount() + " logical");
    OWLProfileReport r = new OWL2DLProfile().checkOntology(o);
    System.out.println("OWL 2 DL: " + (r.isInProfile() ? "yes" : "NO"));
    for (OWLProfileViolation v : r.getViolations()) System.out.println("  violation: " + v);
    if (a.length > 1) { o.axioms().sorted().forEach(ax -> System.out.println("  " + ax)); }
  }
}
