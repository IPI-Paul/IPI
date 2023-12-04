const BubbleMarks = (selection, props) => {
    const {
        worldAtlas: {land, interiors },
        migrants, 
        sizeScale,
        sizeValue
    } = props;           
    
    const { geoNaturalEarth1, geoPath, geoGraticule } = d3;
    const projection = geoNaturalEarth1()
        .scale(window.innerWidth / 2.5 / Math.PI)
        .rotate([0, 0])
        .center([0, window.innerWidth > 1813 ? -25 : -32])
        .translate([window.innerWidth / 2, window.innerHeight / 2])
    ; 
    const path = geoPath(projection);
    const graticule = geoGraticule();

    const groups = selection
        .selectAll('g')
        .data(land.features);

    const groupsEnter = groups
        .enter()
        .append('g')
        .attr('class', 'marks')
    ;

    groupsEnter
        .merge(groups)
        .append('path')
        .attr('class', 'pSphere')
        .attr('d', path({ type: 'Sphere' }))
    ;

    groupsEnter
        .merge(groups)
        .append('path')
        .attr('class', 'pGraticules')
        .attr('d', path(graticule()))
    ;
        
    groups.exit().remove();
    
    groupsEnter
        .data(land.features)
        .append('path')
        .attr('class', 'pLand')
        .attr('d', d => path(d))
        .merge(groups.select('path'))
    ;
    
    groupsEnter
        .append('path')
        .attr('class', 'pInteriors')
        .attr('d', path(interiors))
        .merge(groups.select('path'))
    ;
    
    groupsEnter
        .selectAll('circle')
        .data(migrants)
        .enter()
        .append('circle')
        .attr('cx', d => ([x, y] = projection(d.coords), x))
        .attr('cy', d =>  ([x, y] = projection(d.coords), y))
        .attr('r', d => sizeScale(sizeValue(d)))
        .merge(groups.select('circle'))
    ;
};

globalThis.BubbleMarks = BubbleMarks;
