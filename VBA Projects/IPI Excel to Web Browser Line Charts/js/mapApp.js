(async () => {
    const { select } = d3;
    
    const margin = { top: 30, left: 10}
    const defaults = {width: window.innerWidth - margin.left, height: window.innerHeight - margin.top};
    const width = window.innerWidth - margin.left;
    const height = window.innerHeight - margin.top;
    const dateHistogramSize = 0.2;
    const xValue = d => d['Reported Date'];
    let brushExtent;

    const setBrushExtent = (data) => {
        brushExtent = data;
        svg
            .select('#bubble-container')
            .select('#bubble-map')
            .remove()
        ;
        svg
            .select('#bubble-container')
            .append('g')
                .attr('id', 'bubble-map')
                .call(
                BubbleMap, 
                { 
                    migrants: sourceData,
                    filteredData: brushExtent ? sourceData.filter(d => {
                    const date = xValue(d);
                    return date > data[0] && date < data[1];
                    }) : sourceData,
                    worldAtlas: useWorldAtlas()
                }
        );  
    }

    const render = data => {
        svg
        .append('g')
            .attr('id', 'bubble-container')
        .append('g')
            .attr('id', 'bubble-map')
            .call(
            BubbleMap, 
            { 
                migrants: sourceData,
                filteredData: sourceData,
                worldAtlas: useWorldAtlas()
            }
            ); 
        svg
        .append('g')
            .attr('transform', `translate(0, ${height - (dateHistogramSize * height)})`)
            .call(
            DateHistogram, 
            { 
                migrants: sourceData,
                width,
                height: dateHistogramSize * height,
                setBrushExtent,
                xValue
            }
            ); 
    };
    const svg = select('#root')
        .append('svg')
        .attr('width', width)
        .attr('height', height)
        .attr('viewbox', `0 0 ${width} ${height}`);    
        
    const formatData = d => {
        d['Total Dead and Missing'] = +d['Total Dead and Missing'];
        d['Reported Date'] = new Date(d['Reported Date']);
        d.coords = d['Location Coordinates'].split(',').map(d => +d).reverse();
    };

    sourceData.forEach(formatData);
    sourceData.columns = Object.keys(sourceData[0]);
    render(sourceData);
})();
