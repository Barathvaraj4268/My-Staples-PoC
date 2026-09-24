// 1. Consensus whitespace: competitor shelves with no Staples equivalent, carried by 3+ retailers
MATCH (c:Category) WHERE c.retailer <> 'Staples' AND NOT (c)-[:SAME_AS]->(:Category {retailer:'Staples'})
MATCH (c)-[:SIMILAR_TO]-(d:Category) WHERE d.retailer <> c.retailer AND d.retailer <> 'Staples'
WITH c, collect(DISTINCT d.retailer) + c.retailer AS carriers
WHERE size(carriers) >= 3 RETURN c.path, carriers ORDER BY size(carriers) DESC LIMIT 25;
// (SIMILAR_TO peer edges are written by run_framework.py into outputs/final/graph/peer_edges.csv)

// 2. Aisle whitespace: a missing shelf whose siblings Staples mostly carries
MATCH (c:Category)-[:CHILD_OF]->(p)<-[:CHILD_OF]-(sib)
WHERE c.retailer <> 'Staples' AND NOT (c)-[:SAME_AS]->(:Category {retailer:'Staples'})
WITH c, p, count(sib) AS n_sib,
     sum(CASE WHEN (sib)-[:SAME_AS]->(:Category {retailer:'Staples'}) THEN 1 ELSE 0 END) AS carried
WHERE n_sib >= 3 AND toFloat(carried)/n_sib >= 0.7
RETURN c.retailer, p.path AS aisle, c.name AS missing_shelf, carried, n_sib ORDER BY n_sib DESC LIMIT 25;

// 3. Cannibalisation check: competitor shelves pointing at a Staples CORE shelf
MATCH (c:Category)-[s:SAME_AS]->(t:Category {retailer:'Staples'}) WHERE t.coreness >= 0.7
RETURN c.retailer, c.path, t.path, s.weight, t.coreness ORDER BY t.coreness * s.weight DESC LIMIT 25;

// 4. Adjacency with Graph Data Science: PageRank seeded on Staples leaves
CALL gds.graph.project('cat', 'Category', {CHILD_OF:{orientation:'UNDIRECTED', properties:'weight'},
     CROSS_LISTED_UNDER:{orientation:'UNDIRECTED', properties:'weight'}, SAME_AS:{orientation:'UNDIRECTED', properties:'weight'}});
MATCH (s:Category {retailer:'Staples', is_leaf:true}) WITH collect(s) AS seeds
CALL gds.pageRank.stream('cat', {sourceNodes: seeds, relationshipWeightProperty:'weight'}) YIELD nodeId, score
WITH gds.util.asNode(nodeId) AS n, score WHERE n.retailer <> 'Staples'
RETURN n.retailer, n.path, score ORDER BY score DESC LIMIT 25;
