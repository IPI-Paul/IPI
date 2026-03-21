const Marks = (selection, props) => {
  const {
    nested, 
    colorScale,
    xScale, 
    yScale, 
    xValue, 
    yValue, 
    circleRadius,
    hoveredValue,
    fadeOpacity
  } = props;           
 
  const groups = selection
      .selectAll('g')
      .data(nested);

  const groupsEnter = groups
      .enter()
      .append('g');

  groupsEnter
    .merge(groups);
      
  groups.exit().remove();
  let id = -1;
  let key;
  let id1 = -1;
  let key1;
  let id2 = -1;
  let key2;
  groupsEnter
    .selectAll('circle')
    .data(d => d.values)
    .enter()
    .append('circle')
      .attr('class', 'marks-circle')
    .merge(groups.select('circle'))
      .attr('fill', (d, idx) => {
        idx == 0 && id++; 
        key = nested.map(d => d.key)[id];
        return colorScale(key);
      })
      .attr('opacity', (d, idx) => {
        idx == 0 && id1++; 
        key1 = nested.map(d => d.key)[id1]; 
        return hoveredValue && key1 !== hoveredValue ? fadeOpacity : 1;
      })
      .attr('r', circleRadius * 0.25)
      .attr('cx', d => xScale(xValue(d)))
      .attr('cy', d => yScale(yValue(d)))
        .append('title')          
          .text((d, idx) => {
            idx == 0 && id2++; 
            key2 = nested.map(d => d.key)[id2]; 
            return `${xValue(d)}\n${key2}: ${yValue(d).toFixed(2)}`;
          });
};

globalThis.Marks = Marks;