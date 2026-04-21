#!/usr/bin/env python3
"""RDF operations for the OnlyOffice CLI.

This module isolates RDF graph logic from the main DocumentServerClient
implementation so RDF can evolve independently of DOCX/XLSX/PPTX/PDF code.
"""

from __future__ import annotations

import os
import tempfile
from pathlib import Path
from typing import Any, Dict, Optional, Tuple


class RDFOperations:
    """Encapsulate RDF graph operations using a host client's IO helpers."""

    _RDF_FORMAT_MAP = {
        ".ttl": "turtle",
        ".n3": "n3",
        ".nt": "nt",
        ".nq": "nquads",
        ".jsonld": "json-ld",
        ".json": "json-ld",
        ".rdf": "xml",
        ".xml": "xml",
        ".trig": "trig",
    }

    def __init__(self, host: Any):
        self.host = host

    def rdf_format_from_path(self, file_path: str) -> str:
        """Infer RDF serialisation format from file extension."""
        return self._RDF_FORMAT_MAP.get(Path(file_path).suffix.lower(), "turtle")

    def rdf_safe_save(self, data: str, file_path: str) -> None:
        """Atomically write a serialised RDF string."""
        target = Path(file_path)
        target.parent.mkdir(parents=True, exist_ok=True)
        fd, tmp_path = tempfile.mkstemp(
            prefix=f".{target.name}.",
            dir=str(target.parent),
        )
        try:
            with os.fdopen(fd, "w") as handle:
                handle.write(data)
            os.replace(tmp_path, str(target))
        except Exception:
            try:
                os.unlink(tmp_path)
            except OSError:
                pass
            raise

    def get_graph(self, file_path: str = None) -> Tuple[Optional[Any], Optional[str]]:
        """Load or create an RDF graph. Returns ``(graph, error_string)``."""
        try:
            from rdflib import Graph
        except ImportError:
            return None, "rdflib not installed. pip install rdflib"

        graph = Graph()
        if file_path:
            if not os.path.exists(file_path):
                return None, f"File not found: {file_path}"
            graph.parse(file_path, format=self.rdf_format_from_path(file_path))
        return graph, None

    def create(
        self,
        file_path: str,
        base_uri: str = None,
        format: str = "turtle",
        prefixes: Dict[str, str] = None,
    ) -> Dict[str, Any]:
        """Create a new empty RDF graph file."""
        try:
            from rdflib import Graph, Namespace
            from rdflib.namespace import DCTERMS, FOAF, OWL, RDF, RDFS, SKOS, XSD

            with self.host._file_lock(file_path):
                self.host._snapshot_backup(file_path)
                graph = Graph()
                if base_uri:
                    graph.bind("base", Namespace(base_uri))
                if prefixes:
                    for prefix, uri in prefixes.items():
                        graph.bind(prefix, Namespace(uri))
                for ns_prefix, ns in [
                    ("rdf", RDF),
                    ("rdfs", RDFS),
                    ("owl", OWL),
                    ("xsd", XSD),
                    ("foaf", FOAF),
                    ("dcterms", DCTERMS),
                    ("skos", SKOS),
                ]:
                    graph.bind(ns_prefix, ns)
                self.rdf_safe_save(graph.serialize(format=format), file_path)
            return {
                "success": True,
                "file": file_path,
                "format": format,
                "triples": len(graph),
                "prefixes": {prefix: str(namespace) for prefix, namespace in graph.namespaces()},
            }
        except ImportError:
            return {"success": False, "error": "rdflib not installed. pip install rdflib"}
        except Exception as exc:
            return {"success": False, "error": str(exc)}

    def read(self, file_path: str, limit: int = 100) -> Dict[str, Any]:
        """Read and parse an RDF file, returning triples and stats."""
        try:
            graph, err = self.get_graph(file_path)
            if err:
                return {"success": False, "error": err}
            triples = []
            truncated = False
            for index, (subject, predicate, obj) in enumerate(graph):
                if index >= limit:
                    truncated = True
                    break
                triples.append(
                    {
                        "subject": str(subject),
                        "predicate": str(predicate),
                        "object": str(obj),
                    }
                )
            subjects = {str(subject) for subject in graph.subjects()}
            predicates = {str(predicate) for predicate in graph.predicates()}
            return {
                "success": True,
                "file": file_path,
                "total_triples": len(graph),
                "unique_subjects": len(subjects),
                "unique_predicates": len(predicates),
                "prefixes": {prefix: str(namespace) for prefix, namespace in graph.namespaces()},
                "triples": triples,
                "truncated": truncated,
            }
        except Exception as exc:
            return {"success": False, "error": str(exc)}

    def add(
        self,
        file_path: str,
        subject: str,
        predicate: str,
        object_val: str,
        object_type: str = "uri",
        lang: str = None,
        datatype: str = None,
        format: str = None,
    ) -> Dict[str, Any]:
        """Add a triple to an RDF graph."""
        try:
            from rdflib import BNode, Literal, URIRef

            if lang and datatype:
                return {
                    "success": False,
                    "error": "Cannot specify both --lang and --datatype for a literal",
                }

            with self.host._file_lock(file_path):
                self.host._snapshot_backup(file_path)
                graph, err = self.get_graph(file_path)
                if err:
                    return {"success": False, "error": err}
                fmt = format or self.rdf_format_from_path(file_path)
                subj = URIRef(subject)
                pred = URIRef(predicate)
                if object_type == "literal":
                    dt = URIRef(datatype) if datatype else None
                    obj = Literal(object_val, lang=lang, datatype=dt)
                elif object_type == "bnode":
                    obj = BNode(object_val)
                else:
                    obj = URIRef(object_val)
                before = len(graph)
                graph.add((subj, pred, obj))
                self.rdf_safe_save(graph.serialize(format=fmt), file_path)
            return {
                "success": True,
                "file": file_path,
                "triple": {
                    "subject": str(subj),
                    "predicate": str(pred),
                    "object": str(obj),
                },
                "triples_before": before,
                "triples_after": len(graph),
                "format": fmt,
            }
        except ImportError:
            return {"success": False, "error": "rdflib not installed"}
        except Exception as exc:
            return {"success": False, "error": str(exc)}

    def remove(
        self,
        file_path: str,
        subject: str = None,
        predicate: str = None,
        object_val: str = None,
        object_type: str = "uri",
        lang: str = None,
        datatype: str = None,
        format: str = None,
    ) -> Dict[str, Any]:
        """Remove triples matching a pattern. None values act as wildcards."""
        try:
            from rdflib import BNode, Literal, URIRef

            if lang and datatype:
                return {
                    "success": False,
                    "error": "Cannot specify both --lang and --datatype for a literal",
                }

            with self.host._file_lock(file_path):
                self.host._snapshot_backup(file_path)
                graph, err = self.get_graph(file_path)
                if err:
                    return {"success": False, "error": err}
                fmt = format or self.rdf_format_from_path(file_path)
                subj = URIRef(subject) if subject else None
                pred = URIRef(predicate) if predicate else None
                if object_val is None:
                    obj = None
                elif object_type == "literal":
                    dt = URIRef(datatype) if datatype else None
                    obj = Literal(object_val, lang=lang, datatype=dt)
                elif object_type == "bnode":
                    obj = BNode(object_val)
                else:
                    obj = URIRef(object_val)
                before = len(graph)
                graph.remove((subj, pred, obj))
                self.rdf_safe_save(graph.serialize(format=fmt), file_path)
            return {
                "success": True,
                "file": file_path,
                "triples_before": before,
                "triples_after": len(graph),
                "removed": before - len(graph),
                "format": fmt,
            }
        except ImportError:
            return {"success": False, "error": "rdflib not installed"}
        except Exception as exc:
            return {"success": False, "error": str(exc)}

    def query(self, file_path: str, sparql: str, limit: int = 100) -> Dict[str, Any]:
        """Execute a SPARQL query against an RDF graph."""
        try:
            graph, err = self.get_graph(file_path)
            if err:
                return {"success": False, "error": err}
            results = graph.query(sparql)
            query_type = results.type

            if query_type == "ASK":
                return {
                    "success": True,
                    "file": file_path,
                    "query_type": "ASK",
                    "result": bool(results),
                }

            if query_type in ("CONSTRUCT", "DESCRIBE"):
                triples = []
                truncated = False
                for index, (subject, predicate, obj) in enumerate(results.graph):
                    if index >= limit:
                        truncated = True
                        break
                    triples.append(
                        {
                            "subject": str(subject),
                            "predicate": str(predicate),
                            "object": str(obj),
                        }
                    )
                return {
                    "success": True,
                    "file": file_path,
                    "query_type": query_type,
                    "triples": triples,
                    "result_count": len(triples),
                    "truncated": truncated,
                }

            variables = [str(var) for var in results.vars] if results.vars else []
            rows = []
            truncated = False
            for index, row in enumerate(results):
                if index >= limit:
                    truncated = True
                    break
                rows.append(
                    {
                        str(var): str(row[var]) if row[var] is not None else None
                        for var in results.vars
                    }
                )
            return {
                "success": True,
                "file": file_path,
                "query_type": "SELECT",
                "variables": variables,
                "result_count": len(rows),
                "rows": rows,
                "truncated": truncated,
            }
        except ImportError:
            return {"success": False, "error": "rdflib not installed"}
        except Exception as exc:
            return {"success": False, "error": str(exc)}

    def export(
        self,
        file_path: str,
        output_path: str,
        output_format: str = "turtle",
    ) -> Dict[str, Any]:
        """Convert/export an RDF graph to a different serialisation format."""
        try:
            with self.host._file_lock(output_path):
                self.host._snapshot_backup(output_path)
                graph, err = self.get_graph(file_path)
                if err:
                    return {"success": False, "error": err}
                self.rdf_safe_save(graph.serialize(format=output_format), output_path)
            return {
                "success": True,
                "source": file_path,
                "output": output_path,
                "format": output_format,
                "triples": len(graph),
                "size": os.path.getsize(output_path),
            }
        except ImportError:
            return {"success": False, "error": "rdflib not installed"}
        except Exception as exc:
            return {"success": False, "error": str(exc)}

    def merge(
        self,
        file_path: str,
        other_path: str,
        output_path: str = None,
        format: str = None,
    ) -> Dict[str, Any]:
        """Merge two RDF graphs. Result goes to output_path or overwrites file_path."""
        try:
            abs_a = str(Path(file_path).resolve())
            abs_b = str(Path(other_path).resolve())
            if abs_a == abs_b:
                return {
                    "success": False,
                    "error": "Cannot merge a file with itself — use rdf-export to copy",
                }

            target = output_path or file_path
            with self.host._file_lock(target):
                self.host._snapshot_backup(target)
                graph_a, err = self.get_graph(file_path)
                if err:
                    return {"success": False, "error": err}
                graph_b, err = self.get_graph(other_path)
                if err:
                    return {"success": False, "error": err}
                fmt = format or self.rdf_format_from_path(target)
                before = len(graph_a)
                graph_a += graph_b
                self.rdf_safe_save(graph_a.serialize(format=fmt), target)
            return {
                "success": True,
                "file": target,
                "graph_a_triples": before,
                "graph_b_triples": len(graph_b),
                "merged_triples": len(graph_a),
                "format": fmt,
            }
        except ImportError:
            return {"success": False, "error": "rdflib not installed"}
        except Exception as exc:
            return {"success": False, "error": str(exc)}

    def stats(self, file_path: str) -> Dict[str, Any]:
        """Get detailed statistics about an RDF graph."""
        try:
            from rdflib import Literal as RDFLiteral
            from rdflib.namespace import RDF

            graph, err = self.get_graph(file_path)
            if err:
                return {"success": False, "error": err}
            subjects = {str(subject) for subject in graph.subjects()}
            predicates = {str(predicate) for predicate in graph.predicates()}
            objects_uri = set()
            objects_literal = set()
            for obj in graph.objects():
                if isinstance(obj, RDFLiteral):
                    objects_literal.add(str(obj))
                else:
                    objects_uri.add(str(obj))
            classes = {
                str(obj) for _, _, obj in graph.triples((None, RDF.type, None))
            }
            predicate_frequency: Dict[str, int] = {}
            for _, predicate, _ in graph:
                pred_text = str(predicate)
                predicate_frequency[pred_text] = predicate_frequency.get(pred_text, 0) + 1
            top_predicates = sorted(
                predicate_frequency.items(),
                key=lambda item: item[1],
                reverse=True,
            )[:20]
            return {
                "success": True,
                "file": file_path,
                "total_triples": len(graph),
                "unique_subjects": len(subjects),
                "unique_predicates": len(predicates),
                "unique_objects_uri": len(objects_uri),
                "unique_objects_literal": len(objects_literal),
                "rdf_types": sorted(classes),
                "type_count": len(classes),
                "prefixes": {prefix: str(namespace) for prefix, namespace in graph.namespaces()},
                "top_predicates": [
                    {"predicate": predicate, "count": count}
                    for predicate, count in top_predicates
                ],
            }
        except ImportError:
            return {"success": False, "error": "rdflib not installed"}
        except Exception as exc:
            return {"success": False, "error": str(exc)}

    def namespace(
        self,
        file_path: str,
        prefix: str = None,
        uri: str = None,
        format: str = None,
    ) -> Dict[str, Any]:
        """Add a namespace prefix or list all prefixes."""
        try:
            from rdflib import Namespace

            if prefix and uri:
                with self.host._file_lock(file_path):
                    self.host._snapshot_backup(file_path)
                    graph, err = self.get_graph(file_path)
                    if err:
                        return {"success": False, "error": err}
                    fmt = format or self.rdf_format_from_path(file_path)
                    graph.bind(prefix, Namespace(uri))
                    self.rdf_safe_save(graph.serialize(format=fmt), file_path)
                return {
                    "success": True,
                    "file": file_path,
                    "action": "added",
                    "prefix": prefix,
                    "uri": uri,
                    "all_prefixes": {
                        ns_prefix: str(namespace)
                        for ns_prefix, namespace in graph.namespaces()
                    },
                }

            graph, err = self.get_graph(file_path)
            if err:
                return {"success": False, "error": err}
            return {
                "success": True,
                "file": file_path,
                "prefixes": {
                    ns_prefix: str(namespace)
                    for ns_prefix, namespace in graph.namespaces()
                },
                "prefix_count": len(list(graph.namespaces())),
            }
        except ImportError:
            return {"success": False, "error": "rdflib not installed"}
        except Exception as exc:
            return {"success": False, "error": str(exc)}

    def _extract_validation_violations(self, results_graph: Any) -> list[Dict[str, Any]]:
        from rdflib import URIRef

        shacl = "http://www.w3.org/ns/shacl#"
        violations = []
        for result_node in results_graph.objects(None, URIRef(f"{shacl}result")):
            violation: Dict[str, Any] = {}
            focus_node = next(
                results_graph.objects(result_node, URIRef(f"{shacl}focusNode")),
                None,
            )
            if focus_node:
                violation["focusNode"] = str(focus_node)
            result_path = next(
                results_graph.objects(result_node, URIRef(f"{shacl}resultPath")),
                None,
            )
            if result_path:
                violation["resultPath"] = str(result_path)
            message = next(
                results_graph.objects(result_node, URIRef(f"{shacl}resultMessage")),
                None,
            )
            if message:
                violation["message"] = str(message)
            severity = next(
                results_graph.objects(result_node, URIRef(f"{shacl}resultSeverity")),
                None,
            )
            if severity:
                violation["severity"] = str(severity).split("#")[-1]
            constraint = next(
                results_graph.objects(
                    result_node,
                    URIRef(f"{shacl}sourceConstraintComponent"),
                ),
                None,
            )
            if constraint:
                violation["constraint"] = str(constraint).split("#")[-1]
            violations.append(violation)
        return violations

    def validate(self, file_path: str, shapes_path: str) -> Dict[str, Any]:
        """Validate an RDF graph against SHACL shapes."""
        try:
            from pyshacl import validate as shacl_validate

            graph, err = self.get_graph(file_path)
            if err:
                return {"success": False, "error": err}
            shapes_graph, err = self.get_graph(shapes_path)
            if err:
                return {"success": False, "error": f"Shapes file error: {err}"}

            conforms, results_graph, results_text = shacl_validate(
                graph,
                shacl_graph=shapes_graph,
                inference="rdfs",
            )

            violations = []
            violations_parse_error = None
            try:
                violations = self._extract_validation_violations(results_graph)
            except Exception as exc:
                violations_parse_error = str(exc)

            return {
                "success": True,
                "file": file_path,
                "shapes_file": shapes_path,
                "conforms": conforms,
                "violation_count": len(violations),
                "violations": violations,
                "violations_parse_error": violations_parse_error,
                "results_text": results_text[:5000],
            }
        except ImportError:
            return {"success": False, "error": "pyshacl not installed. pip install pyshacl"}
        except Exception as exc:
            return {"success": False, "error": str(exc)}
