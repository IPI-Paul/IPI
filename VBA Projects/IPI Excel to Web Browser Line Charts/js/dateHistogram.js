
const { 
    scaleTime, scaleLinear, timeFormat, extent, bin, timeMonths, sum,  
    format, brushX, select 
} = d3;

const margin = { top: 0, right: 30, bottom: 5, left: 60 };
const xAxisLabelOffset = 50;
const yAxisLabelOffset = 30;
const tickOffset = 10;
const barOffset = 3;
const xAxisTickFormat = timeFormat('%b-%Y');
const tooltipFormat = format(',');
let xScale;
let yScale;
let binnedData;

const DateHistogram = (selection, props) => {
    const {migrants, width, height, setBrushExtent, xValue} = props;

    const innerHeight = height - margin.top - margin.bottom;
    const innerWidth = width - margin.left - margin.right;

    const xAxisLabel = 'Reported Date';

    const yValue = d => d['Total Dead and Missing'];
    const yAxisLabel = 'Total Dead and Missing';

    let idx = 0;

    if (!xScale) {
        xScale = scaleTime()
        .domain(extent(migrants, xValue))
        .range([0, innerWidth])
        .nice();
    }

    const [start, stop] = xScale.domain();

    if(!binnedData){
        binnedData = bin()
        .value(xValue)
        .domain(xScale.domain())
        .thresholds(timeMonths(start, stop))
        (migrants)
        .map(array => ({
            y: sum(array, yValue),
            x0: array.x0,
            x1: array.x1
        }));
    }

    if(!yScale){
        yScale = scaleLinear()
        .domain([0, max(binnedData, d => d.y)])
        .range([innerHeight, 0]);
    }

    const group = selection
        .append('g');

    const groupEnter = group
        .enter()
        .append('g')
    ;

    groupEnter
        .merge(group)
        .append('rect')
        .attr('width', width)
        .attr('height', height)
        .attr('class', 'bgWhite')
    ;

    groupEnter
        .merge(group)
        .append('g')
        .attr('transform', `translate(${margin.left}, ${margin.top})`)
        .attr('class', 'multi')
    ;

    groupEnter
        .merge(group)
        .call(
        AxisBottom,
        {
            xScale,
            innerHeight,
            tickFormat: xAxisTickFormat,
            tickOffset
        }
        )
    ;

    groupEnter
        .merge(group)
        .call(
        AxisLeft,
        {
            yScale,
            innerWidth,
            tickOffset
        }
        )
    ;

    groupEnter
        .merge(group)
        .append('text')
        .attr('x', innerWidth / 2)
        .attr('y', innerHeight + xAxisLabelOffset)
        .attr('style', 'text-anchor: middle;')
        .attr('class', 'axis-label-multi')
        .text(xAxisLabel)
    ;

    groupEnter
        .merge(group)
        .append('text')
        .attr('transform', `translate(${yAxisLabelOffset - 10}, ${innerHeight / 2}) rotate(-90)`)
        .attr('style', 'text-anchor: middle;')
        .attr('class', 'axis-label-multi')
        .text(yAxisLabel)
    ;

    groupEnter
        .merge(group)
        .call(
        HistogramMarks,
        {
            binnedData,
            xScale,
            yScale,
            tooltipFormat,
            innerHeight,
            barOffset
        }
        )
    ;
        
    group.exit().remove();

    selection
        .append('g')
        .attr('id', 'brushRef')
    ;

    select('#brushRef')
        .call(
        brushX().extent([[margin.left, 0], [innerWidth + margin.left, innerHeight]])
            .on('brush end', (event, d) => {
            const selection = event.selection && [
                event.selection[0] - margin.left, 
                event.selection[1] - margin.left
            ];
            if(event.type === 'end' || idx > 3) {
                idx = 0
                setBrushExtent(selection && selection.map(xScale.invert));
            } else {
                idx++
            }
            })
        )
    ;
}
