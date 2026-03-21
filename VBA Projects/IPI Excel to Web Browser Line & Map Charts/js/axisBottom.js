const AxisBottom = (selection, props) => {
    const {xScale, innerHeight, tickFormat, tickOffset} = props;

    selection
        .select('.multi')
        .selectAll('.bottomTick')
        .data(xScale.ticks())
        .enter()
        .append('g')
        .attr('transform', d => `translate(${xScale(d)}, 0)`)
        .attr('class', 'bottomTick')
        .append('line')
        .attr('y2', innerHeight)
        .select(function () { return this.parentNode; })
        .append('text')
        .attr('dy', '0.71em')
        .attr('y', innerHeight + (tickOffset ? tickOffset : 3))
        .attr('style', 'text-anchor: middle;')
        .text(d => tickFormat(d))
    ;
}
