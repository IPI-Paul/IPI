
(async () => {
  const { 
      select, csv, scaleLinear, scaleOrdinal, scaleTime, extent, axisLeft, axisBottom, 
      line, curveBasis, group, schemeCategory10, descending, timeFormat, curveNatural
  } = d3;

  let hoveredValue;

  const defaults = {width: 1150, height: 500};
  const client = {width: document.body.clientWidth, height: document.body.clientHeight};
  const topMargin = 45;
  const width = client.width < defaults.width ? client.width : defaults.width;
  const height = (client.height < defaults.height ? client.height : defaults.height) - topMargin;
  
  const widthRatio = (input) => (
      height < defaults.height - topMargin 
      ? ((input / innerWidth) * width) - ((input * .3) * (width / defaults.width))
      : input
  );

  const heightRatio = (input) => (
      height < defaults.height - topMargin
      ? (input / (defaults.height - topMargin)) * height
      : (input / height) * (defaults.height - topMargin)
  );

  const margin = {top: 90, right: 185, bottom: 80, left: 110};
  const inner = {
      width: width - margin.left - margin.right,
      height: height - margin.top - margin.bottom,
      top: 20,
      right: 170,
      bottom: 80,
      left: 80,
      x: 30,
      y: 25,
      radius: 10
  };
    
  let xColumn;
  let yColumn;
  const filtered = () => {
    let dta;
    if (xColumn !== "All" && yColumn !== "All") {
      dta = sourceData.filter(d => d.city === xColumn ||  d.city === yColumn);
      dta.columns = Object.keys(dta[0]);
    } else {
      dta = sourceData
    }
    svg.selectAll('*').remove()
    return dta;
  };

  const onXColumnClicked = column => {
    xColumn = column;
    render(filtered());
  };

  const onYColumnClicked = column => {
    yColumn = column;
    render(filtered());
  };

  const render = data => {
      const columns = (items) => {
      const obj = data.columns.map(d => 
          ({
          value: d,
          text: d.toProperCase().replace('stamp', ''),
          type: ['timestamp'].includes(d) 
              ? 'ordinal' 
              : 'quantitative'
          })
      );
      return [obj.find(d => d.value === items[0]), obj.find(d => d.value === items[1])]
      };

      const [xType, yType] = columns(['timestamp', 'temperature']);
      const xValue = d => d[xType.value];
      const xAxisLabel = xType.text;
      const yValue = d => d[yType.value];
      const yAxisLabel = yType.text;
      const chartTitle = `A Week of Temperature Around the World`;
      const colorValue = d => d.city;
      const fadeOpacity = 0.2;

      if (!xColumn) {
        xColumn = "All";
        yColumn = "All";
      }
      
      const menus = select('#menu')
        .selectAll('.menus')
        .data([null])
        .enter()
        .append('div')
          .attr('class', 'menus')
          .attr('style', 'width: ' + defaults.width + 'px;');

      let menuValues = ["All"];
      Array.from(
        group(data, colorValue), 
        ([key, values]) => key
      ).map(key => menuValues.push(key));

      menus
        .call(dropdownMenu, {
          options: menuValues,
          onOptionClicked: onYColumnClicked, 
          selectedOption: yColumn,
          id: 'y-menu'
        });

      menus.selectAll('text').data([null]).enter().append('text').text(' ')

      menus
        .call(dropdownMenu, {
          options: menuValues,
          onOptionClicked: onXColumnClicked, 
          selectedOption: xColumn,
          id: 'x-menu'
        });

      const xScale = (xType.type === 'ordinal' 
      ? scaleTime()
      : scaleLinear()
      )
      .domain(extent(data, xValue))  
      .range([0, inner.width])      
      .nice();

      const yScale = (yType.type === 'ordinal' 
      ? scaleTime()
      : scaleLinear()  
      )
      .domain(extent(data, yValue))
      .range([inner.height, 0])    
      .nice();

      const colorScale = scaleOrdinal(schemeCategory10);
  
      const g = svg
      .append('g')
          .attr('transform', `translate(${margin.left}, ${margin.top})`)
          .attr('class', 'line-chart');

      const xAxis = axisBottom(xScale)
      .tickSize(-inner.height)
      // xTickFormat
      .tickPadding(15);

      const yAxis = axisLeft(yScale)
      .tickSize(-inner.width)
      // yTickFormat
      .tickPadding(10);

      const yAxisG = g.append('g')
      .call(yAxis);

      yAxisG
      .selectAll('.domain')
          .remove();

      yAxisG
      .append('text')
          .attr('class', 'axis-label')
          .attr('y', -inner.left)
          .attr('x', -inner.height / 2)
          .attr('fill', 'black')
          .attr('transform', `rotate(-90)`)
          .attr('text-anchor', 'middle')
      .text(yAxisLabel);

      const xAxisG = g.append('g')
      .call(xAxis)
          .attr('transform', `translate(0, ${inner.height})`);

      xAxisG.select('.domain').remove();

      xAxisG.append('text')
      .attr('class', 'axis-label')
      .attr('y', inner.bottom)
      .attr('x', inner.width / 2)
      .attr('fill', 'black')
      .text(xAxisLabel);

      const lineGenerator = line()
      .x(d => xScale(xValue(d)))
      .y(d => yScale(yValue(d)))
      .curve(curveNatural);

      const lastYValue = d => yValue(d.values[d.values.length - 1]);

      const nested = Array.from(
          group(data, colorValue), 
          ([key, values]) => ({key, values})
      )
      .sort((a, b) => 
          descending(lastYValue(a), lastYValue(b))
      );

      colorScale.domain(nested.map(d => d.key));
      
      g
      .selectAll('.multi-line-path')
      .data(nested)
      .enter()
      .append('path')
        .attr('class', 'multi-line-path')
        .attr('d', d => lineGenerator(d.values))
        .attr('stroke', d => colorScale(d.key))
        .attr('opacity', d => hoveredValue && d.key !== hoveredValue ? fadeOpacity : 1)
      ;
      
      g.append('text')
      .attr('class', 'title')
      .attr('y', -inner.top)
      .attr('x', inner.width / 2)
      .text(chartTitle);

      svg
      .append('g')
          .attr('transform', `translate(${width - inner.right}, ${heightRatio(height / 3)})`)
          .call(
              colorLegend, 
              { 
                  colorScale, 
                  circleRadius: heightRatio(inner.radius), 
                  spacing: heightRatio(inner.y),
                  textOffset: heightRatio(inner.y - (inner.radius * 0.6)), 
                  onHover: (value) => {hoveredValue = value; svg.selectAll('*').remove(); render(data);},  
                  hoveredValue: hoveredValue,
                  fadeOpacity: fadeOpacity 
              }
          ); 

      // LineMarkers
  };
  const svg = select('#root')
      .append('svg')
      .attr('width', width)
      .attr('height', height)
      .attr('viewbox', `0 0 ${defaults.width} ${defaults.height}`);    
      
  const formatData = d => {
      d.temperature = +d.temperature;
      d.timestamp = new Date(d.timestamp);
      d.Month = new Date(d.Month);
      d.Activity = +d.Activity;
      d['Reported Date'] = new Date(d['Reported Date']);
      d['Total Dead and Missing'] = +d['Total Dead and Missing'];
  };

  sourceData.forEach(formatData);
  sourceData.columns = Object.keys(sourceData[0]);
  render(sourceData);
})();
