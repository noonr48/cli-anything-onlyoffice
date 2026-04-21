import io
import json
import os
import sys
import tempfile
import types
import unittest
from contextlib import redirect_stdout
from unittest import mock

from rdflib import BNode, Graph, Literal, Namespace, URIRef

from cli_anything.onlyoffice.core import cli as cli_module
from cli_anything.onlyoffice.utils.docserver import get_client


class OnlyOfficeRDFTests(unittest.TestCase):
    def setUp(self):
        self.client = get_client()
        self.tmpdir = tempfile.TemporaryDirectory(prefix="oo-rdf-test-")
        self.base = self.tmpdir.name

    def tearDown(self):
        self.tmpdir.cleanup()

    def _path(self, name: str) -> str:
        return os.path.join(self.base, name)

    def _write_turtle(self, path: str, content: str) -> None:
        with open(path, "w", encoding="utf-8") as handle:
            handle.write(content)

    def test_rdf_create_binds_default_and_custom_prefixes(self):
        path = self._path("graph.ttl")

        result = self.client.rdf_create(
            path,
            base_uri="http://example.org/base/",
            prefixes={"ex": "http://example.org/"},
        )

        self.assertTrue(result["success"])
        self.assertEqual(result["format"], "turtle")
        self.assertEqual(result["triples"], 0)
        self.assertIn("ex", result["prefixes"])
        self.assertIn("rdf", result["prefixes"])
        self.assertEqual(result["prefixes"]["base"], "http://example.org/base/")

    def test_rdf_add_read_and_remove_support_literals(self):
        path = self._path("people.ttl")
        self.assertTrue(self.client.rdf_create(path)["success"])

        knows = self.client.rdf_add(
            path,
            "http://example.org/Alice",
            "http://xmlns.com/foaf/0.1/knows",
            "http://example.org/Bob",
        )
        self.assertTrue(knows["success"])

        label = self.client.rdf_add(
            path,
            "http://example.org/Alice",
            "http://www.w3.org/2000/01/rdf-schema#label",
            "Alice",
            object_type="literal",
            lang="en",
        )
        self.assertTrue(label["success"])

        read = self.client.rdf_read(path, limit=10)
        self.assertTrue(read["success"])
        self.assertEqual(read["total_triples"], 2)

        invalid = self.client.rdf_add(
            path,
            "http://example.org/Alice",
            "http://www.w3.org/2000/01/rdf-schema#comment",
            "Bad",
            object_type="literal",
            lang="en",
            datatype="http://www.w3.org/2001/XMLSchema#string",
        )
        self.assertFalse(invalid["success"])
        self.assertIn("Cannot specify both", invalid["error"])

        removed = self.client.rdf_remove(
            path,
            predicate="http://www.w3.org/2000/01/rdf-schema#label",
            object_val="Alice",
            object_type="literal",
            lang="en",
        )
        self.assertTrue(removed["success"])
        self.assertEqual(removed["removed"], 1)

    def test_rdf_query_supports_select_ask_construct_and_describe(self):
        path = self._path("query.ttl")
        self._write_turtle(
            path,
            """@prefix ex: <http://example.org/> .
@prefix foaf: <http://xmlns.com/foaf/0.1/> .
ex:Alice a foaf:Person ;
    foaf:name "Alice" ;
    foaf:knows ex:Bob .
ex:Bob a foaf:Person ;
    foaf:name "Bob" .
""",
        )

        select_result = self.client.rdf_query(
            path,
            "SELECT ?name WHERE { <http://example.org/Alice> <http://xmlns.com/foaf/0.1/name> ?name }",
        )
        self.assertTrue(select_result["success"])
        self.assertEqual(select_result["query_type"], "SELECT")
        self.assertEqual(select_result["rows"][0]["name"], "Alice")

        ask_result = self.client.rdf_query(
            path,
            "ASK { <http://example.org/Alice> <http://xmlns.com/foaf/0.1/knows> <http://example.org/Bob> }",
        )
        self.assertTrue(ask_result["success"])
        self.assertEqual(ask_result["query_type"], "ASK")
        self.assertTrue(ask_result["result"])

        construct_result = self.client.rdf_query(
            path,
            "CONSTRUCT { ?s <http://example.org/linkedTo> ?o } WHERE { ?s <http://xmlns.com/foaf/0.1/knows> ?o }",
        )
        self.assertTrue(construct_result["success"])
        self.assertEqual(construct_result["query_type"], "CONSTRUCT")
        self.assertGreaterEqual(construct_result["result_count"], 1)

        describe_result = self.client.rdf_query(
            path,
            "DESCRIBE <http://example.org/Alice>",
        )
        self.assertTrue(describe_result["success"])
        self.assertEqual(describe_result["query_type"], "DESCRIBE")
        self.assertGreaterEqual(describe_result["result_count"], 1)

    def test_rdf_export_merge_and_namespace_are_supported(self):
        a_path = self._path("a.ttl")
        b_path = self._path("b.ttl")
        merged_path = self._path("merged.ttl")
        export_path = self._path("merged.rdf")

        self._write_turtle(
            a_path,
            "@prefix ex: <http://example.org/> . ex:Alice ex:knows ex:Bob .\n",
        )
        self._write_turtle(
            b_path,
            "@prefix ex: <http://example.org/> . ex:Bob ex:knows ex:Carol .\n",
        )
        self._write_turtle(export_path, "stale")

        merged = self.client.rdf_merge(a_path, b_path, output_path=merged_path)
        self.assertTrue(merged["success"])
        self.assertEqual(merged["merged_triples"], 2)

        exported = self.client.rdf_export(merged_path, export_path, output_format="xml")
        self.assertTrue(exported["success"])
        self.assertEqual(exported["format"], "xml")
        self.assertGreater(exported["size"], 0)

        namespace_added = self.client.rdf_namespace(
            merged_path,
            prefix="schema",
            uri="http://schema.org/",
            format="turtle",
        )
        self.assertTrue(namespace_added["success"])
        self.assertEqual(namespace_added["action"], "added")
        self.assertIn("schema", namespace_added["all_prefixes"])

        namespace_list = self.client.rdf_namespace(merged_path)
        self.assertTrue(namespace_list["success"])
        self.assertIn("schema", namespace_list["prefixes"])

    def test_rdf_stats_reports_types_and_top_predicates(self):
        path = self._path("stats.ttl")
        self._write_turtle(
            path,
            """@prefix ex: <http://example.org/> .
@prefix foaf: <http://xmlns.com/foaf/0.1/> .
ex:Alice a foaf:Person ;
    foaf:name "Alice" ;
    foaf:knows ex:Bob .
ex:Bob a foaf:Person ;
    foaf:name "Bob" .
""",
        )

        result = self.client.rdf_stats(path)

        self.assertTrue(result["success"])
        self.assertEqual(result["total_triples"], 5)
        self.assertIn("http://xmlns.com/foaf/0.1/Person", result["rdf_types"])
        predicates = {item["predicate"] for item in result["top_predicates"]}
        self.assertIn("http://xmlns.com/foaf/0.1/name", predicates)

    def test_rdf_validate_returns_structured_violations(self):
        data_path = self._path("data.ttl")
        shapes_path = self._path("shapes.ttl")
        self._write_turtle(data_path, "@prefix ex: <http://example.org/> . ex:A ex:p ex:B .\n")
        self._write_turtle(shapes_path, "@prefix sh: <http://www.w3.org/ns/shacl#> . [] a sh:NodeShape .\n")

        sh = Namespace("http://www.w3.org/ns/shacl#")
        report_graph = Graph()
        report = BNode()
        result = BNode()
        report_graph.add((report, sh.result, result))
        report_graph.add((result, sh.focusNode, URIRef("http://example.org/A")))
        report_graph.add((result, sh.resultPath, URIRef("http://example.org/p")))
        report_graph.add((result, sh.resultMessage, Literal("Missing value")))
        report_graph.add((result, sh.resultSeverity, sh.Violation))
        report_graph.add((result, sh.sourceConstraintComponent, sh.MinCountConstraintComponent))

        fake_module = types.ModuleType("pyshacl")
        fake_module.validate = lambda *args, **kwargs: (False, report_graph, "validation report")

        with mock.patch.dict(sys.modules, {"pyshacl": fake_module}):
            result = self.client.rdf_validate(data_path, shapes_path)

        self.assertTrue(result["success"])
        self.assertFalse(result["conforms"])
        self.assertEqual(result["violation_count"], 1)
        self.assertEqual(result["violations"][0]["focusNode"], "http://example.org/A")
        self.assertEqual(result["violations"][0]["severity"], "Violation")
        self.assertIsNone(result["violations_parse_error"])

    def test_rdf_validate_surfaces_violation_parse_errors(self):
        data_path = self._path("data_parse.ttl")
        shapes_path = self._path("shapes_parse.ttl")
        self._write_turtle(data_path, "@prefix ex: <http://example.org/> . ex:A ex:p ex:B .\n")
        self._write_turtle(shapes_path, "@prefix sh: <http://www.w3.org/ns/shacl#> . [] a sh:NodeShape .\n")

        class BrokenResultsGraph:
            def objects(self, *args, **kwargs):
                raise RuntimeError("boom")

        fake_module = types.ModuleType("pyshacl")
        fake_module.validate = lambda *args, **kwargs: (False, BrokenResultsGraph(), "validation report")

        with mock.patch.dict(sys.modules, {"pyshacl": fake_module}):
            result = self.client.rdf_validate(data_path, shapes_path)

        self.assertTrue(result["success"])
        self.assertEqual(result["violation_count"], 0)
        self.assertIn("boom", result["violations_parse_error"])

    def test_cli_rdf_query_dispatches_via_rdf_handler(self):
        path = self._path("dispatch.ttl")
        self._write_turtle(path, "@prefix ex: <http://example.org/> . ex:A ex:p ex:B .\n")
        stdout = io.StringIO()

        with mock.patch.object(
            cli_module.doc_server,
            "rdf_query",
            return_value={"success": True, "file": path, "query_type": "SELECT", "rows": []},
        ) as rdf_query:
            with mock.patch(
                "sys.argv",
                [
                    "cli-anything-onlyoffice",
                    "rdf-query",
                    path,
                    "SELECT * WHERE { ?s ?p ?o }",
                    "--limit",
                    "5",
                    "--json",
                ],
            ):
                with redirect_stdout(stdout):
                    cli_module.main()

        payload = json.loads(stdout.getvalue())
        self.assertTrue(payload["success"])
        rdf_query.assert_called_once_with(path, "SELECT * WHERE { ?s ?p ?o }", limit=5)

    def test_cli_rdf_validate_dispatches_via_rdf_handler(self):
        data_path = self._path("dispatch_data.ttl")
        shapes_path = self._path("dispatch_shapes.ttl")
        self._write_turtle(data_path, "@prefix ex: <http://example.org/> . ex:A ex:p ex:B .\n")
        self._write_turtle(shapes_path, "@prefix sh: <http://www.w3.org/ns/shacl#> . [] a sh:NodeShape .\n")
        stdout = io.StringIO()

        with mock.patch.object(
            cli_module.doc_server,
            "rdf_validate",
            return_value={"success": True, "file": data_path, "conforms": True},
        ) as rdf_validate:
            with mock.patch(
                "sys.argv",
                [
                    "cli-anything-onlyoffice",
                    "rdf-validate",
                    data_path,
                    shapes_path,
                    "--json",
                ],
            ):
                with redirect_stdout(stdout):
                    cli_module.main()

        payload = json.loads(stdout.getvalue())
        self.assertTrue(payload["success"])
        rdf_validate.assert_called_once_with(data_path, shapes_path)
