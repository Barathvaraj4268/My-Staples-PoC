// ---- Staples marketplace PoC: six-retailer category knowledge graph ----
// Copy nodes.csv + edges.csv into Neo4j's import folder, then run:
CREATE CONSTRAINT cat_uid IF NOT EXISTS FOR (c:Category) REQUIRE c.uid IS UNIQUE;
LOAD CSV WITH HEADERS FROM 'file:///nodes.csv' AS r
MERGE (c:Category {uid: r.uid})
SET c.retailer = r.retailer, c.name = r.name, c.path = r.path, c.level = toInteger(r.level),
    c.items = toFloat(r.items), c.n_leaves = toInteger(r.n_leaves), c.is_leaf = (r.is_leaf = 'True'),
    c.coreness = toFloat(r.coreness)
MERGE (rt:Retailer {name: r.retailer})
MERGE (c)-[:SOLD_BY]->(rt);
// one LOAD per relationship type (no APOC needed)
LOAD CSV WITH HEADERS FROM 'file:///edges.csv' AS r WITH r WHERE r.rel = 'CHILD_OF'
MATCH (a:Category {uid: r.source}), (b:Category {uid: r.target}) MERGE (a)-[x:CHILD_OF]->(b) SET x.weight = toFloat(r.weight);
LOAD CSV WITH HEADERS FROM 'file:///edges.csv' AS r WITH r WHERE r.rel = 'CROSS_LISTED_UNDER'
MATCH (a:Category {uid: r.source}), (b:Category {uid: r.target}) MERGE (a)-[x:CROSS_LISTED_UNDER]->(b) SET x.weight = toFloat(r.weight);
LOAD CSV WITH HEADERS FROM 'file:///edges.csv' AS r WITH r WHERE r.rel = 'SAME_AS'
MATCH (a:Category {uid: r.source}), (b:Category {uid: r.target}) MERGE (a)-[x:SAME_AS]->(b) SET x.weight = toFloat(r.weight);
