const { scaleSqrt, max } = d3;

let sizeScale;

const BubbleMap = (selection, props) => {
    const { migrants, filteredData, worldAtlas } = props;
    const sizeValue = d => d['Total Dead and Missing'];
    const maxRadius = 15;

    if(!sizeScale) {
    sizeScale = scaleSqrt()
        .domain([0, max(migrants, sizeValue)])
        .range([0, maxRadius]);
    }     

    selection
        .append('g')
        .call(
            BubbleMarks, 
            { 
            worldAtlas,
            migrants: filteredData,
            sizeScale,
            sizeValue,
            }
        ); 
};
