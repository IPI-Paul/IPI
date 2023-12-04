const HistogramMarks = (selection, props) => {
    const { binnedData, xScale, yScale, tooltipFormat, innerHeight, barOffset } = props;

    selection
        .select('.multi')
        .selectAll('histoMark')
        .data(binnedData)
        .enter()
        .append('rect')
        .attr('x', d => xScale(d.x0))
        .attr('y', d => yScale(d.y))
        .attr('width', d => xScale(d.x1) - xScale(d.x0) - barOffset)
        .attr('height', d => innerHeight - yScale(d.y))
        .attr('class', 'histoMark')
        .append('title')
        .text(d => tooltipFormat(d.y))
    ;
};

