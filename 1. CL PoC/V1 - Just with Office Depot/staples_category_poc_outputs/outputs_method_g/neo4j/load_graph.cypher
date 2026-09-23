// ---- Staples marketplace PoC: category knowledge graph ----
// Copy nodes.csv + edges.csv into Neo4j's import folder, then run:
CREATE CONSTRAINT cat_uid IF NOT EXISTS FOR (c:Category) REQUIRE c.uid IS UNIQUE;

LOAD CSV WITH HEADERS FROM 'file:///nodes.csv' AS r
MERGE (c:Category {uid: r.uid})
SET c.retailer = r.retailer, c.name = r.name, c.path = r.path, c.level = toInteger(r.level),
    c.items = toFloat(r.items), c.is_leaf = (r.is_leaf = 'True'), c.coreness = toFloat(r.coreness)
MERGE (rt:Retailer {name: r.retailer})
MERGE (c)-[:SOLD_BY]->(rt);

LOAD CSV WITH HEADERS FROM 'file:///edges.csv' AS r
MATCH (a:Category {uid: r.source}), (b:Category {uid: r.target})
CALL apoc.merge.relationship(a, r.rel, {}, {weight: toFloat(r.weight)}, b, {}) YIELD rel
RETURN count(rel);
// (without APOC: run one LOAD CSV per relationship type with a WHERE r.rel = '...' filter)
