function loadExcelFile() {
    const fileName = 'a1-cars.csv';
    document.getElementById('loading').style.display = 'block';
    document.getElementById('error').style.display = 'none';
    document.getElementById('visualizations').style.display = 'none';

    fetch(fileName)
        .then(response => {
            if (!response.ok) {
                throw new Error('File not found or inaccessible');
            }
            return response.arrayBuffer();
        })
        .then(data => {
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheet = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheet];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            const headers = jsonData[0];
            const rawData = jsonData.slice(1).map(row => {
                const obj = {};
                headers.forEach((header, i) => {
                    obj[header] = row[i];
                });
                return obj;
            });

            const processedData = processAndCleanData(rawData);
            if (processedData.length === 0) {
                throw new Error('No valid data after processing');
            }

            document.getElementById('loading').style.display = 'none';
            document.getElementById('visualizations').style.display = 'grid';
            createVisualizations(processedData);
        })
        .catch(err => {
            document.getElementById('loading').style.display = 'none';
            document.getElementById('error').style.display = 'block';
            document.getElementById('error').textContent = 'Error loading file: ' + err.message;
        });
}

function processAndCleanData(data) {
    return data.filter(d => {
        return d.MPG && d.Horsepower && d.Horsepower !== 'NA' && d.Weight && d.Acceleration && d['Model Year'] && d.Origin;
    }).map(d => ({
        Car: d.Car || 'Unknown',
        Manufacturer: d.Manufacturer || 'Unknown',
        MPG: +d.MPG,
        Cylinders: +d.Cylinders,
        Displacement: +d.Displacement || 0,
        Horsepower: +d.Horsepower,
        Weight: +d.Weight,
        Acceleration: +d.Acceleration,
        ModelYear: +d['Model Year'],
        Origin: d.Origin
    }));
}

function createVisualizations(data) {
    const container = document.querySelector('.vis');
    let width = container.clientWidth - 20;
    const maxWidth = 800;
    if (width > maxWidth) width = maxWidth;
    const height = 450;
    const margin = { top: 30, right: 150, bottom: 60, left: 60 };
    const plotWidth = width - margin.left - margin.right;
    const plotHeight = height - margin.top - margin.bottom;

    const color = d3.scaleOrdinal()
        .domain(['American', 'European', 'Japanese'])
        .range(['red', 'blue', 'green']);

    const tooltip = d3.select('body').append('div')
        .attr('class', 'tooltip')
        .style('opacity', 0);

    const visStates = {
        scatter1: {
            selectedData: data.slice(),
            transform: d3.zoomIdentity,
            zoom: d3.zoom().scaleExtent([1, 10]).on('zoom', (event) => zoomed(event, 'scatter1'))
        },
        barchart: {
            selectedData: data.slice(),
            transform: d3.zoomIdentity,
            zoom: d3.zoom().scaleExtent([1, 10]).on('zoom', (event) => zoomed(event, 'barchart')),
            brush: d3.brushX()
                .extent([[0, 0], [plotWidth, plotHeight]])
                .on('brush end', (event) => brushed(event, 'barchart'))
        },
        boxplot: {
            selectedData: data.slice(),
            transform: d3.zoomIdentity,
            zoom: d3.zoom().scaleExtent([1, 10]).on('zoom', (event) => zoomed(event, 'boxplot'))
        },
        linechart: {
            selectedData: data.slice(),
            transform: d3.zoomIdentity,
            zoom: d3.zoom().scaleExtent([1, 10]).on('zoom', (event) => zoomed(event, 'linechart'))
        }
    };

    let globalOriginFilter = 'all';

    const scatter1Svg = d3.select('#scatter1')
        .append('svg')
        .attr('width', width)
        .attr('height', height)
        .attr('viewBox', `0 0 ${width} ${height}`)
        .attr('preserveAspectRatio', 'xMidYMid meet');

    const scatter1 = scatter1Svg.append('g')
        .attr('transform', `translate(${margin.left},${margin.top})`);

    const x1 = d3.scaleLinear()
        .domain([0, d3.max(data, d => d.MPG) * 1.1])
        .range([0, plotWidth]);

    const y1 = d3.scaleLinear()
        .domain([0, d3.max(data, d => d.Horsepower) * 1.1])
        .range([plotHeight, 0]);

    const x1Axis = scatter1.append('g')
        .attr('class', 'x-axis')
        .attr('transform', `translate(0,${plotHeight})`)
        .call(d3.axisBottom(x1))
        .append('text')
        .attr('x', plotWidth / 2)
        .attr('y', 40)
        .attr('fill', 'black')
        .style('font-size', '16px')
        .text('MPG');

    const y1Axis = scatter1.append('g')
        .attr('class', 'y-axis')
        .call(d3.axisLeft(y1))
        .append('text')
        .attr('transform', 'rotate(-90)')
        .attr('x', -plotHeight / 2)
        .attr('y', -40)
        .attr('fill', 'black')
        .style('font-size', '16px')
        .text('Horsepower');

    scatter1Svg.call(visStates.scatter1.zoom);

    scatter1Svg.append('foreignObject')
        .attr('x', width - margin.right + 20)
        .attr('y', margin.top + 70)
        .attr('width', 120)
        .attr('height', 30)
        .html(`
            <button class="reset-button" data-vis="scatter1">Reset</button>
        `);

    const scatterLegend = scatter1Svg.append('g')
        .attr('class', 'legend')
        .attr('transform', `translate(${width - margin.right + 20}, ${margin.top})`);

    color.domain().forEach((origin, i) => {
        const legendItem = scatterLegend.append('g')
            .attr('class', 'legend-item')
            .attr('transform', `translate(0, ${i * 20})`);
        legendItem.append('circle')
            .attr('class', 'legend-dot')
            .attr('cx', 5)
            .attr('cy', 5)
            .attr('r', 5)
            .attr('fill', color(origin));
        legendItem.append('text')
            .attr('x', 15)
            .attr('y', 9)
            .text(origin);
    });

    const barSvg = d3.select('#barchart')
        .append('svg')
        .attr('width', width)
        .attr('height', height)
        .attr('viewBox', `0 0 ${width} ${height}`)
        .attr('preserveAspectRatio', 'xMidYMid meet');

    const bar = barSvg.append('g')
        .attr('transform', `translate(${margin.left},${margin.top})`);

    const years = [...new Set(data.map(d => d.ModelYear))].sort();
    const cylinders = [...new Set(data.map(d => d.Cylinders))].sort();
    const colorBar = d3.scaleOrdinal()
        .domain(cylinders)
        .range(d3.schemeCategory10);

    const xBar = d3.scaleBand()
        .domain(years)
        .range([0, plotWidth])
        .padding(0.1);

    const yBar = d3.scaleLinear()
        .range([plotHeight, 0]);

    const xBarAxis = bar.append('g')
        .attr('class', 'x-axis')
        .attr('transform', `translate(0,${plotHeight})`)
        .call(d3.axisBottom(xBar))
        .append('text')
        .attr('x', plotWidth / 2)
        .attr('y', 40)
        .attr('fill', 'black')
        .style('font-size', '16px')
        .text('Model Year');

    const yBarAxis = bar.append('g')
        .attr('class', 'y-axis')
        .call(d3.axisLeft(yBar))
        .append('text')
        .attr('transform', 'rotate(-90)')
        .attr('x', -plotHeight / 2)
        .attr('y', -40)
        .attr('fill', 'black')
        .style('font-size', '16px')
        .text('Count');

    bar.append('g')
        .attr('class', 'brush')
        .call(visStates.barchart.brush);

    barSvg.call(visStates.barchart.zoom);

    barSvg.append('foreignObject')
        .attr('x', width - margin.right + 20)
        .attr('y', margin.top + 90)
        .attr('width', 120)
        .attr('height', 30)
        .html(`
            <button class="reset-button" data-vis="barchart">Reset</button>
        `);

    const barLegend = barSvg.append('g')
        .attr('class', 'legend')
        .attr('transform', `translate(${width - margin.right + 20}, ${margin.top})`);

    cylinders.forEach((cyl, i) => {
        const legendItem = barLegend.append('g')
            .attr('class', 'legend-item')
            .attr('transform', `translate(0, ${i * 20})`);
        legendItem.append('rect')
            .attr('class', 'legend-rect')
            .attr('x', 0)
            .attr('y', 0)
            .attr('width', 10)
            .attr('height', 10)
            .attr('fill', colorBar(cyl));
        legendItem.append('text')
            .attr('x', 15)
            .attr('y', 9)
            .text(`${cyl} Cylinders`);
    });

    const boxSvg = d3.select('#boxplot')
        .append('svg')
        .attr('width', width)
        .attr('height', height)
        .attr('viewBox', `0 0 ${width} ${height}`)
        .attr('preserveAspectRatio', 'xMidYMid meet');

    const box = boxSvg.append('g')
        .attr('transform', `translate(${margin.left},${margin.top})`);

    const origins = ['American', 'European', 'Japanese'];
    const xBox = d3.scaleBand()
        .domain(origins)
        .range([0, plotWidth])
        .padding(0.4);

    const yBox = d3.scaleLinear()
        .domain([0, d3.max(data, d => d.MPG) * 1.1])
        .range([plotHeight, 0]);

    const xBoxAxis = box.append('g')
        .attr('class', 'x-axis')
        .attr('transform', `translate(0,${plotHeight})`)
        .call(d3.axisBottom(xBox))
        .append('text')
        .attr('x', plotWidth / 2)
        .attr('y', 40)
        .attr('fill', 'black')
        .style('font-size', '16px')
        .text('Origin');

    const yBoxAxis = box.append('g')
        .attr('class', 'y-axis')
        .call(d3.axisLeft(yBox))
        .append('text')
        .attr('transform', 'rotate(-90)')
        .attr('x', -plotHeight / 2)
        .attr('y', -40)
        .attr('fill', 'black')
        .style('font-size', '16px')
        .text('MPG');

    boxSvg.call(visStates.boxplot.zoom);

    boxSvg.append('foreignObject')
        .attr('x', width - margin.right + 20)
        .attr('y', margin.top + 70)
        .attr('width', 120)
        .attr('height', 30)
        .html(`
            <button class="reset-button" data-vis="boxplot">Reset</button>
        `);

    const boxLegend = boxSvg.append('g')
        .attr('class', 'legend')
        .attr('transform', `translate(${width - margin.right + 20}, ${margin.top})`);

    origins.forEach((origin, i) => {
        const legendItem = boxLegend.append('g')
            .attr('class', 'legend-item')
            .attr('transform', `translate(0, ${i * 20})`);
        legendItem.append('rect')
            .attr('class', 'legend-rect')
            .attr('x', 0)
            .attr('y', 0)
            .attr('width', 10)
            .attr('height', 10)
            .attr('fill', color(origin));
        legendItem.append('text')
            .attr('x', 15)
            .attr('y', 9)
            .text(origin);
    });

    const lineSvg = d3.select('#linechart')
        .append('svg')
        .attr('width', width)
        .attr('height', height)
        .attr('viewBox', `0 0 ${width} ${height}`)
        .attr('preserveAspectRatio', 'xMidYMid meet');

    const line = lineSvg.append('g')
        .attr('transform', `translate(${margin.left},${margin.top})`);

    const xLine = d3.scaleLinear()
        .domain(d3.extent(years))
        .range([0, plotWidth]);

    const yLine = d3.scaleLinear()
        .domain([0, d3.max(data, d => d.MPG) * 1.1])
        .range([plotHeight, 0]);

    const xLineAxis = line.append('g')
        .attr('class', 'x-axis')
        .attr('transform', `translate(0,${plotHeight})`)
        .call(d3.axisBottom(xLine).tickFormat(d3.format('d')))
        .append('text')
        .attr('x', plotWidth / 2)
        .attr('y', 40)
        .attr('fill', 'black')
        .style('font-size', '16px')
        .text('Model Year');

    const yLineAxis = line.append('g')
        .attr('class', 'y-axis')
        .call(d3.axisLeft(yLine))
        .append('text')
        .attr('transform', 'rotate(-90)')
        .attr('x', -plotHeight / 2)
        .attr('y', -40)
        .attr('fill', 'black')
        .style('font-size', '16px')
        .text('Average MPG');

    lineSvg.call(visStates.linechart.zoom);

    lineSvg.append('foreignObject')
        .attr('x', width - margin.right + 20)
        .attr('y', margin.top + 70)
        .attr('width', 120)
        .attr('height', 30)
        .html(`
            <button class="reset-button" data-vis="linechart">Reset</button>
        `);

    const lineLegend = lineSvg.append('g')
        .attr('class', 'legend')
        .attr('transform', `translate(${width - margin.right + 20}, ${margin.top})`);

    origins.forEach((origin, i) => {
        const legendItem = lineLegend.append('g')
            .attr('class', 'legend-item')
            .attr('transform', `translate(0, ${i * 20})`);
        legendItem.append('line')
            .attr('class', 'legend-line')
            .attr('x1', 0)
            .attr('x2', 20)
            .attr('y1', 5)
            .attr('y2', 5)
            .attr('stroke', color(origin))
            .attr('stroke-width', 2);
        legendItem.append('text')
            .attr('x', 25)
            .attr('y', 9)
            .text(origin);
    });

    const globalControls = d3.select('body')
        .insert('div', ':first-child')
        .attr('class', 'global-controls');

    globalControls.append('label')
        .attr('for', 'global-origin-filter')
        .style('font-size', '12px')
        .style('margin-right', '10px')
        .text('Filter by Origin:');

    globalControls.append('select')
        .attr('id', 'global-origin-filter')
        .attr('class', 'origin-filter')
        .html(`
            <option value="all">All</option>
            <option value="American">American</option>
            <option value="European">European</option>
            <option value="Japanese">Japanese</option>
        `);

    globalControls.append('button')
        .attr('id', 'global-reset-button')
        .attr('class', 'reset-button')
        .style('margin-left', '10px')
        .text('Reset All');

    function getBarData(selectedData) {
        const stackedData = years.map(year => {
            const yearData = selectedData.filter(d => d.ModelYear === year);
            const counts = cylinders.reduce((acc, cyl) => {
                acc[cyl] = yearData.filter(d => d.Cylinders === cyl).length;
                return acc;
            }, {});
            return { year, ...counts };
        });
        return d3.stack().keys(cylinders)(stackedData);
    }

    function getBoxData(selectedData) {
        return origins.map(origin => {
            const values = selectedData.filter(d => d.Origin === origin).map(d => d.MPG).sort(d3.ascending);
            const q1 = d3.quantile(values, 0.25);
            const median = d3.quantile(values, 0.5);
            const q3 = d3.quantile(values, 0.75);
            const iqr = q3 - q1;
            const min = Math.max(d3.min(values) || 0, q1 - 1.5 * iqr);
            const max = Math.min(d3.max(values) || d3.max(data, d => d.MPG), q3 + 1.5 * iqr);
            const outliers = values.filter(v => v < min || v > max);
            return { origin, q1, median, q3, min, max, outliers };
        });
    }

    function getLineData(selectedData) {
        const lineData = origins.map(origin => {
            const yearlyData = years.map(year => {
                const yearOriginData = selectedData.filter(d => d.Origin === origin && d.ModelYear === year);
                const avgMPG = yearOriginData.length > 0 ? d3.mean(yearOriginData, d => d.MPG) : null;
                return { year, avgMPG, origin };
            });
            return { origin, values: yearlyData.filter(d => d.avgMPG !== null) };
        });
        return lineData;
    }

    function updateVisData() {
        Object.keys(visStates).forEach(visId => {
            if (globalOriginFilter === 'all') {
                visStates[visId].selectedData = data.slice();
            } else {
                visStates[visId].selectedData = data.filter(d => d.Origin === globalOriginFilter).slice();
            }
        });
    }

    d3.select('#global-origin-filter').on('change', function() {
        globalOriginFilter = this.value;
        updateVisData();
        d3.select('#barchart .brush').call(visStates.barchart.brush.move, null);
        Object.keys(visStates).forEach(visId => {
            updateVisualization(visId);
        });
    });

    d3.selectAll('.reset-button').on('click', function() {
        const visId = this.getAttribute('data-vis');
        visStates[visId].selectedData = globalOriginFilter === 'all' ? data.slice() : data.filter(d => d.Origin === globalOriginFilter).slice();
        visStates[visId].transform = d3.zoomIdentity;
        const svg = d3.select(`#${visId} svg`);
        svg.call(visStates[visId].zoom.transform, d3.zoomIdentity);
        if (visId === 'barchart') {
            d3.select('#barchart .brush').call(visStates.barchart.brush.move, null);
        }
        updateVisualization(visId);
    });

    d3.select('#global-reset-button').on('click', function() {
        globalOriginFilter = 'all';
        d3.select('#global-origin-filter').property('value', 'all');
        Object.keys(visStates).forEach(visId => {
            visStates[visId].selectedData = data.slice();
            visStates[visId].transform = d3.zoomIdentity;
            const svg = d3.select(`#${visId} svg`);
            svg.call(visStates[visId].zoom.transform, d3.zoomIdentity);
            if (visId === 'barchart') {
                d3.select('#barchart .brush').call(visStates.barchart.brush.move, null);
            }
            updateVisualization(visId);
        });
    });

    function resizeVisualizations() {
        const newWidth = container.clientWidth - 20;
        width = newWidth > maxWidth ? maxWidth : newWidth;
        const newPlotWidth = width - margin.left - margin.right;

        x1.range([0, newPlotWidth]);
        xBar.range([0, newPlotWidth]);
        xBox.range([0, newPlotWidth]);
        xLine.range([0, newPlotWidth]);

        [scatter1Svg, barSvg, boxSvg, lineSvg].forEach(svg => {
            svg.attr('width', width)
               .attr('viewBox', `0 0 ${width} ${height}`);
            svg.selectAll('foreignObject')
                .attr('x', width - margin.right + 20);
        });

        scatter1Svg.select('.legend').attr('transform', `translate(${width - margin.right + 20}, ${margin.top})`);
        barSvg.select('.legend').attr('transform', `translate(${width - margin.right + 20}, ${margin.top})`);
        boxSvg.select('.legend').attr('transform', `translate(${width - margin.right + 20}, ${margin.top})`);
        lineSvg.select('.legend').attr('transform', `translate(${width - margin.right + 20}, ${margin.top})`);

        visStates.barchart.brush.extent([[0, 0], [newPlotWidth, plotHeight]]);
        bar.select('.brush').call(visStates.barchart.brush);

        Object.keys(visStates).forEach(visId => updateVisualization(visId));
    }

    window.addEventListener('resize', resizeVisualizations);

    Object.keys(visStates).forEach(visId => updateVisualization(visId));

    function updateVisualization(visId) {
        const state = visStates[visId];
        const selectedData = state.selectedData;
        const transform = state.transform;

        if (visId === 'scatter1') {
            const newX1 = transform.rescaleX(x1);
            const newY1 = transform.rescaleY(y1);

            scatter1.selectAll('.dot')
                .data(selectedData)
                .join('circle')
                .attr('class', 'dot')
                .attr('cx', d => transform.applyX(x1(d.MPG)))
                .attr('cy', d => transform.applyY(y1(d.Horsepower)))
                .attr('r', 5)
                .attr('fill', d => color(d.Origin))
                .attr('opacity', 0.7)
                .on('mouseover', (event, d) => {
                    tooltip.transition().duration(200).style('opacity', 0.9);
                    tooltip.html(`Car: ${d.Car}<br>MPG: ${d.MPG}<br>Horsepower: ${d.Horsepower}<br>Origin: ${d.Origin}`)
                        .style('left', (event.pageX + 10) + 'px')
                        .style('top', (event.pageY - 28) + 'px');
                })
                .on('mouseout', () => tooltip.transition().duration(500).style('opacity', 0));

            scatter1.select('.x-axis').call(d3.axisBottom(newX1));
            scatter1.select('.y-axis').call(d3.axisLeft(newY1));
        } else if (visId === 'barchart') {
            const stackedData = getBarData(selectedData);
            yBar.domain([-5, d3.max(stackedData[stackedData.length - 1], d => d[1]) || 1]);

            const xBarZoomed = d3.scaleLinear()
                .domain([0, years.length])
                .range(transform.rescaleX(d3.scaleLinear().domain([0, years.length]).range([0, plotWidth])).range());

            bar.selectAll('g.layer')
                .data(stackedData)
                .join('g')
                .attr('class', 'layer')
                .attr('fill', d => colorBar(d.key))
                .selectAll('rect')
                .data(d => d)
                .join('rect')
                .attr('x', d => {
                    const index = years.indexOf(d.data.year);
                    return xBarZoomed(index);
                })
                .attr('y', d => transform.applyY(yBar(d[1])))
                .attr('height', d => transform.applyY(yBar(d[0])) - transform.applyY(yBar(d[1])))
                .attr('width', xBar.bandwidth() * transform.k)
                .on('mouseover', (event, d) => {
                    tooltip.transition().duration(200).style('opacity', 0.9);
                    const cyl = event.target.parentNode.__data__.key;
                    tooltip.html(`Year: ${d.data.year}<br>Cylinders: ${cyl}<br>Count: ${d[1] - d[0]}`)
                        .style('left', (event.pageX + 10) + 'px')
                        .style('top', (event.pageY - 28) + 'px');
                })
                .on('mouseout', () => tooltip.transition().duration(500).style('opacity', 0));

            bar.select('.x-axis').call(d3.axisBottom(xBarZoomed).tickValues(years.map((_, i) => i)).tickFormat(i => years[i]));
            bar.select('.y-axis').call(d3.axisLeft(transform.rescaleY(yBar)));
        } else if (visId === 'boxplot') {
            const boxData = getBoxData(selectedData);
            const xBoxZoomed = d3.scaleLinear()
                .domain([0, origins.length])
                .range(transform.rescaleX(d3.scaleLinear().domain([0, origins.length]).range([0, plotWidth])).range());

            const boxGroups = box.selectAll('.box-group')
                .data(boxData)
                .join('g')
                .attr('class', 'box-group')
                .attr('transform', d => {
                    const index = origins.indexOf(d.origin);
                    return `translate(${xBoxZoomed(index)},0)`;
                });

            const boxWidth = xBox.bandwidth() * transform.k;

            boxGroups.selectAll('.range-line')
                .data(d => [d])
                .join('line')
                .attr('class', 'range-line')
                .attr('x1', boxWidth / 2)
                .attr('x2', boxWidth / 2)
                .attr('y1', d => transform.applyY(yBox(d.min)))
                .attr('y2', d => transform.applyY(yBox(d.max)))
                .attr('stroke', 'black');

            boxGroups.selectAll('.box')
                .data(d => [d])
                .join('rect')
                .attr('class', 'box')
                .attr('x', 0)
                .attr('y', d => transform.applyY(yBox(d.q3)))
                .attr('width', boxWidth)
                .attr('height', d => transform.applyY(yBox(d.q1)) - transform.applyY(yBox(d.q3)))
                .attr('fill', d => color(d.origin))
                .attr('opacity', 0.7)
                .on('mouseover', (event, d) => {
                    tooltip.transition().duration(200).style('opacity', 0.9);
                    tooltip.html(`Origin: ${d.origin}<br>Median MPG: ${d.median.toFixed(1)}<br>Q1: ${d.q1.toFixed(1)}<br>Q3: ${d.q3.toFixed(1)}<br>Min: ${d.min.toFixed(1)}<br>Max: ${d.max.toFixed(1)}`)
                        .style('left', (event.pageX + 10) + 'px')
                        .style('top', (event.pageY - 28) + 'px');
                })
                .on('mouseout', () => tooltip.transition().duration(500).style('opacity', 0));

            boxGroups.selectAll('.median-line')
                .data(d => [d])
                .join('line')
                .attr('class', 'median-line')
                .attr('x1', 0)
                .attr('x2', boxWidth)
                .attr('y1', d => transform.applyY(yBox(d.median)))
                .attr('y2', d => transform.applyY(yBox(d.median)))
                .attr('stroke', 'black')
                .attr('stroke-width', 2);

            boxGroups.selectAll('.whisker-min')
                .data(d => [d])
                .join('line')
                .attr('class', 'whisker-min')
                .attr('x1', boxWidth / 4)
                .attr('x2', boxWidth * 3 / 4)
                .attr('y1', d => transform.applyY(yBox(d.min)))
                .attr('y2', d => transform.applyY(yBox(d.min)))
                .attr('stroke', 'black');

            boxGroups.selectAll('.whisker-max')
                .data(d => [d])
                .join('line')
                .attr('class', 'whisker-max')
                .attr('x1', boxWidth / 4)
                .attr('x2', boxWidth * 3 / 4)
                .attr('y1', d => transform.applyY(yBox(d.max)))
                .attr('y2', d => transform.applyY(yBox(d.max)))
                .attr('stroke', 'black');

            boxGroups.selectAll('.outlier')
                .data(d => d.outliers.map(val => ({ origin: d.origin, value: val })))
                .join('circle')
                .attr('class', 'outlier')
                .attr('cx', boxWidth / 2)
                .attr('cy', d => transform.applyY(yBox(d.value)))
                .attr('r', 3)
                .attr('fill', d => color(d.origin))
                .on('mouseover', (event, d) => {
                    tooltip.transition().duration(200).style('opacity', 0.9);
                    tooltip.html(`Origin: ${d.origin}<br>MPG: ${d.value.toFixed(1)}`)
                        .style('left', (event.pageX + 10) + 'px')
                        .style('top', (event.pageY - 28) + 'px');
                })
                .on('mouseout', () => tooltip.transition().duration(500).style('opacity', 0));

            box.select('.x-axis').call(d3.axisBottom(xBoxZoomed).tickValues(origins.map((_, i) => i)).tickFormat(i => origins[i]));
            box.select('.y-axis').call(d3.axisLeft(transform.rescaleY(yBox)));
        } else if (visId === 'linechart') {
            const lineData = getLineData(selectedData);
            const lineGenerator = d3.line()
                .x(d => transform.applyX(xLine(d.year)))
                .y(d => transform.applyY(yLine(d.avgMPG)))
                .defined(d => d.avgMPG !== null);

            line.selectAll('.line')
                .data(lineData)
                .join('path')
                .attr('class', 'line')
                .attr('d', d => lineGenerator(d.values))
                .attr('fill', 'none')
                .attr('stroke', d => color(d.origin))
                .attr('stroke-width', 2);

            line.selectAll('.line-points')
                .data(lineData)
                .join('g')
                .attr('class', 'line-points')
                .selectAll('circle')
                .data(d => d.values)
                .join('circle')
                .attr('cx', d => transform.applyX(xLine(d.year)))
                .attr('cy', d => transform.applyY(yLine(d.avgMPG)))
                .attr('r', 4)
                .attr('fill', d => color(d.origin))
                .on('mouseover', (event, d) => {
                    tooltip.transition().duration(200).style('opacity', 0.9);
                    tooltip.html(`Origin: ${d.origin}<br>Year: ${d.year}<br>Avg MPG: ${d.avgMPG.toFixed(1)}`)
                        .style('left', (event.pageX + 10) + 'px')
                        .style('top', (event.pageY - 28) + 'px');
                })
                .on('mouseout', () => tooltip.transition().duration(500).style('opacity', 0));

            line.select('.x-axis').call(d3.axisBottom(transform.rescaleX(xLine)).tickFormat(d3.format('d')));
            line.select('.y-axis').call(d3.axisLeft(transform.rescaleY(yLine)));
        }
    }

    function zoomed(event, visId) {
        visStates[visId].transform = event.transform;
        updateVisualization(visId);
    }

    function brushed(event, visId) {
        if (visId !== 'barchart') return;

        if (!event.selection) {
            updateVisData();
            Object.keys(visStates).forEach(v => updateVisualization(v));
            return;
        }

        const [x0, x1] = event.selection;
        const yearIndices = years.map((_, i) => i);
        const xBarZoomed = d3.scaleLinear()
            .domain([0, years.length])
            .range(visStates.barchart.transform.rescaleX(d3.scaleLinear().domain([0, years.length]).range([0, plotWidth])).range());
        const selectedYears = years.filter((year, i) => {
            const x = xBarZoomed(i);
            return x >= x0 && x <= x1;
        });

        let filteredData = data.filter(d => selectedYears.includes(d.ModelYear));
        if (globalOriginFilter !== 'all') {
            filteredData = filteredData.filter(d => d.Origin === globalOriginFilter);
        }

        Object.keys(visStates).forEach(v => {
            visStates[v].selectedData = filteredData.slice();
            visStates[v].transform = d3.zoomIdentity;
            d3.select(`#${v} svg`).call(visStates[v].zoom.transform, d3.zoomIdentity);
            updateVisualization(v);
        });
    }
}

loadExcelFile();