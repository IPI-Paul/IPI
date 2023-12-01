
    const colorLegend = (selection, props) => {
        const { colorScale, circleRadius, spacing, textOffset, onHover, hoveredValue, fadeOpacity } = props;
          
        const groups = selection
          .selectAll('g')
          .data(colorScale.domain());
  
        const groupsEnter = groups
          .enter()
          .append('g')
        ;
        
        groupsEnter
          .merge(groups)
            .attr('transform', (d, i) => `translate(0, ${i * spacing})`)
            .on('mouseenter', (_, d) => {if (d !== hoveredValue) {onHover(d)}})
            .on('mouseout', (_, d) => {onHover(null)})
        ;
          
        groups.exit().remove();
  
        groupsEnter
          .append('circle')
            .attr('class', 'legend-circle')
          .merge(groups.select('circle'))
            .attr('fill', colorScale)
            .attr('opacity', d => hoveredValue && d !== hoveredValue ? fadeOpacity : 1)
            .attr('r', circleRadius)
        ;
  
        groupsEnter
          .append('text')
            .attr('class', 'multi-line-legend')
          .merge(groups.select('text'))        
            .text(d => d)
            .attr('dy', '0.32em')
            .attr('x', textOffset)
        ;
      };
      
      globalThis.colorLegend = colorLegend;
  