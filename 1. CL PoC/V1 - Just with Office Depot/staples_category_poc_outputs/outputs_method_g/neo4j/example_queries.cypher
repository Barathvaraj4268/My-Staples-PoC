// ---- Example questions the graph answers directly ----

// 1. Whitespace: competitor shelves with NO equivalent at Staples, whose sibling shelves
//    mostly DO have one (adjacent by construction - "Staples already sells the rest of this aisle")
MATCH (c:Category {retailer:'OfficeDepot'})-[:CHILD_OF]->(p)<-[:CHILD_OF]-(sib)
WHERE NOT (c)-[:SAME_AS]->(:Category {retailer:'Staples'})
WITH c, p, count(sib) AS n_sib,
     sum(CASE WHEN (sib)-[:SAME_AS]->(:Category {retailer:'Staples'}) THEN 1 ELSE 0 END) AS carried
WHERE n_sib >= 3 AND toFloat(carried)/n_sib >= 0.7
RETURN p.path AS aisle, c.name AS missing_shelf, c.items AS competitor_items, carried, n_sib
ORDER BY competitor_items DESC LIMIT 25;

// 2. Consensus whitespace across competitors (once 6 retailers are loaded):
MATCH (c:Category)-[:SAME_AS]-(d:Category)
WHERE c.retailer <> 'Staples' AND d.retailer <> 'Staples'
  AND NOT (c)-[:SAME_AS]-(:Category {retailer:'Staples'})
RETURN c.name, collect(DISTINCT d.retailer) AS also_carried_by, size(collect(DISTINCT d.retailer)) AS peers
ORDER BY peers DESC LIMIT 25;

// 3. Cannibalization check: does a candidate shelf point at a Staples CORE shelf?
MATCH (c:Category {retailer:'OfficeDepot'})-[s:SAME_AS]->(t:Category {retailer:'Staples'})
WHERE t.coreness >= 0.7
RETURN c.path, t.path, s.weight, t.coreness ORDER BY t.coreness * s.weight DESC LIMIT 25;

// 4. Adjacency with Graph Data Science: PageRank seeded on Staples shelves
CALL gds.graph.project('cat', 'Category', {CHILD_OF:{orientation:'UNDIRECTED', properties:'weight'},
     CROSS_LISTED_UNDER:{orientation:'UNDIRECTED', properties:'weight'},
     SAME_AS:{orientation:'UNDIRECTED', properties:'weight'}});
MATCH (s:Category {retailer:'Staples', is_leaf:true}) WITH collect(s) AS seeds
CALL gds.pageRank.stream('cat', {sourceNodes: seeds, relationshipWeightProperty:'weight'})
YIELD nodeId, score
WITH gds.util.asNode(nodeId) AS n, score WHERE n.retailer = 'OfficeDepot'
RETURN n.path, score ORDER BY score DESC LIMIT 25;

// 5. Staples taxonomy health: shelves that only hold items through cross-listing
MATCH (p:Category {retailer:'Staples'})<-[:CROSS_LISTED_UNDER]-(c)
WHERE NOT (p)<-[:CHILD_OF]-()
RETURN p.path, count(c) AS cross_listed_children ORDER BY cross_listed_children DESC LIMIT 25;
