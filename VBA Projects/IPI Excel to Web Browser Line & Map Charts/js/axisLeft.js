const AxisLeft = (selection, props) => {
    const {yScale, innerWidth, tickOffset} = props;
    selection
        .select('.multi')
        .selectAll('.leftTick')
        .data(yScale.ticks())
        .enter()
        .append('g')
        .attr('transform', d => `translate(0, ${yScale(d)})`)
        .attr('class', 'leftTick')
        .append('line')
        .attr('x2', innerWidth)
        .select(function () { return this.parentNode; })
        .append('text')
        .attr('dy', '0.32em')
        .attr('x', -(tickOffset ? tickOffset : 3))
        .attr('style', 'text-anchor: end;')
        .text(d => d)
    ;
}
