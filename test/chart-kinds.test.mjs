import { test } from 'node:test';
import assert from 'node:assert/strict';
import { DocxDocument } from '../dist/document.js';
import { DocxEditor } from '../dist/editor.js';
import { WORD_NS, REL_NS, OFFICE_REL_NS } from '../dist/xml.js';
import { errorBarRanges, radarPoint, trendlinePoints } from '../dist/chart.js';

const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const encoder = new TextEncoder();

const num = (values) => `<c:numRef><c:f>Sheet1!B2</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${values.map((value, index) => value === null ? '' : `<c:pt idx="${index}"><c:v>${value}</c:v></c:pt>`).join('')}</c:numCache></c:numRef>`;
const str = (values) => `<c:strRef><c:f>Sheet1!A2</c:f><c:strCache><c:ptCount val="${values.length}"/>${values.map((value, index) => `<c:pt idx="${index}"><c:v>${value}</c:v></c:pt>`).join('')}</c:strCache></c:strRef>`;
const ser = (index, name, inner) => `<c:ser><c:idx val="${index}"/><c:order val="${index}"/><c:tx>${str([name])}</c:tx>${inner}</c:ser>`;
const valAx = (id, cross, pos, extra = '') => `<c:valAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${pos}"/>${extra}<c:crossAx val="${cross}"/></c:valAx>`;
const catAx = (id, cross, pos = 'b', deleted = 0) => `<c:catAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="${deleted}"/><c:axPos val="${pos}"/><c:crossAx val="${cross}"/></c:catAx>`;

function chartOf(plotArea) {
  const doc = DocxDocument.create();
  const drawing = `<w:drawing><wp:inline><wp:extent cx="4572000" cy="2743200"/><wp:docPr id="1"/><a:graphic><a:graphicData uri="${C}"><c:chart r:id="rIdChart"/></a:graphicData></a:graphic></wp:inline></w:drawing>`;
  doc.setPartXml(doc.mainDocumentPath, `<w:document xmlns:w="${WORD_NS}" xmlns:wp="${WP}" xmlns:a="${A}" xmlns:c="${C}" xmlns:r="${OFFICE_REL_NS}"><w:body><w:p><w:r>${drawing}</w:r></w:p><w:sectPr/></w:body></w:document>`);
  doc.addPart('word/_rels/document.xml.rels', encoder.encode(`<Relationships xmlns="${REL_NS}"><Relationship Id="rIdChart" Type="${OFFICE_REL_NS}/chart" Target="charts/chart1.xml"/></Relationships>`), 'application/vnd.openxmlformats-package.relationships+xml');
  doc.addPart('word/charts/chart1.xml', encoder.encode(`<c:chartSpace xmlns:c="${C}" xmlns:a="${A}"><c:chart><c:plotArea>${plotArea}</c:plotArea></c:chart></c:chartSpace>`), 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml');
  return doc.getShapes()[0].chart;
}

test('a combo chart keeps each series\' own type and puts the second value axis on the secondary side', () => {
  const chart = chartOf(`
    <c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>
      ${ser(0, '收入', `<c:cat>${str(['Q1', 'Q2', 'Q3'])}</c:cat><c:val>${num([120, 150, 170])}</c:val>`)}
      <c:gapWidth val="219"/><c:axId val="101"/><c:axId val="102"/></c:barChart>
    <c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>
      ${ser(1, '增长率', `<c:cat>${str(['Q1', 'Q2', 'Q3'])}</c:cat><c:val>${num([0.05, 0.25, 0.13])}</c:val>`)}
      <c:marker val="1"/><c:axId val="201"/><c:axId val="202"/></c:lineChart>
    ${catAx(101, 102)}${valAx(102, 101, 'l', '<c:majorGridlines/>')}${valAx(202, 201, 'r')}${catAx(201, 202, 'b', 1)}`);
  assert.equal(chart.kind, 'bar');
  assert.deepEqual(chart.series.map((series) => [series.name, series.type, series.axis]), [
    ['收入', 'bar', 'primary'], ['增长率', 'line', 'secondary'],
  ]);
  assert.deepEqual(chart.axes.secondaryValue, { visible: true });
  assert.equal(chart.axes.value.majorGridlines, true);
  // 自动配色按 c:idx：第二条序列是 accent2，不是按组内序号再从 accent1 数起。
  assert.notEqual(chart.series[0].fill.color, chart.series[1].fill.color);
});

test('radar, bubble, stock and 3D charts read instead of falling back to a placeholder', () => {
  const radar = chartOf(`<c:radarChart><c:radarStyle val="filled"/><c:varyColors val="0"/>
    ${ser(0, 'A', `<c:cat>${str(['x', 'y', 'z'])}</c:cat><c:val>${num([1, 2, 3])}</c:val>`)}<c:axId val="1"/><c:axId val="2"/></c:radarChart>
    ${catAx(1, 2)}${valAx(2, 1, 'l')}`);
  assert.equal(radar.kind, 'radar');
  assert.equal(radar.radarStyle, 'filled');
  assert.deepEqual(radar.categories, ['x', 'y', 'z']);

  const bubble = chartOf(`<c:bubbleChart><c:varyColors val="0"/>
    ${ser(0, 'B', `<c:xVal>${num([1, 2, 3])}</c:xVal><c:yVal>${num([5, 6, 7])}</c:yVal><c:bubbleSize>${num([10, 20, 30])}</c:bubbleSize><c:bubble3D val="0"/>`)}
    <c:bubbleScale val="100"/><c:showNegBubbles val="0"/><c:axId val="1"/><c:axId val="2"/></c:bubbleChart>
    ${valAx(1, 2, 'b')}${valAx(2, 1, 'l')}`);
  assert.equal(bubble.kind, 'bubble');
  assert.deepEqual(bubble.series[0].xValues, [1, 2, 3]);
  assert.deepEqual(bubble.series[0].values, [5, 6, 7]);
  assert.deepEqual(bubble.series[0].bubbleSizes, [10, 20, 30]);
  assert.equal(bubble.series[0].axis, 'primary', '散点 / 气泡的竖轴（axPos l）才是数值轴');

  const stock = chartOf(`<c:stockChart>
    ${['Open', 'High', 'Low', 'Close'].map((name, index) => ser(index, name, `<c:spPr><a:ln w="19050"><a:noFill/></a:ln></c:spPr><c:cat>${str(['d1', 'd2'])}</c:cat><c:val>${num([[10, 12], [13, 14], [9, 10], [12, 11]][index])}</c:val>`)).join('')}
    <c:hiLowLines/><c:upDownBars><c:gapWidth val="150"/><c:upBars/><c:downBars/></c:upDownBars><c:axId val="1"/><c:axId val="2"/></c:stockChart>
    ${catAx(1, 2)}${valAx(2, 1, 'l')}`);
  assert.equal(stock.kind, 'stock');
  assert.deepEqual(stock.stock, { hiLowLines: true, upDownBars: true });
  assert.equal(stock.series.length, 4);

  const bar3d = chartOf(`<c:bar3DChart><c:barDir val="bar"/><c:grouping val="clustered"/>
    ${ser(0, 'S', `<c:cat>${str(['a'])}</c:cat><c:val>${num([1])}</c:val>`)}<c:shape val="box"/><c:axId val="1"/><c:axId val="2"/><c:axId val="0"/></c:bar3DChart>
    ${catAx(1, 2, 'l')}${valAx(2, 1, 'b')}`);
  assert.equal(bar3d.kind, 'bar');
  assert.equal(bar3d.threeD, true);
  assert.equal(bar3d.barDirection, 'bar');

  const surface = chartOf(`<c:surfaceChart>${ser(0, 'S', `<c:val>${num([1])}</c:val>`)}<c:axId val="1"/><c:axId val="2"/><c:axId val="3"/></c:surfaceChart>`);
  assert.equal(surface.kind, 'unsupported', '曲面图（等高线）不画');
});

test('trendlines and error bars read with their parameters', () => {
  const chart = chartOf(`<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>
    ${ser(0, 'D', `<c:trendline><c:name>拟合</c:name><c:trendlineType val="poly"/><c:order val="3"/><c:dispRSqr val="0"/><c:dispEq val="0"/></c:trendline>
      <c:trendline><c:trendlineType val="movingAvg"/><c:period val="4"/></c:trendline>
      <c:errBars><c:errDir val="y"/><c:errBarType val="both"/><c:errValType val="percentage"/><c:noEndCap val="0"/><c:val val="5"/></c:errBars>
      <c:errBars><c:errDir val="x"/><c:errBarType val="plus"/><c:errValType val="cust"/><c:noEndCap val="0"/><c:plus>${num([0.1, 0.2])}</c:plus></c:errBars>
      <c:xVal>${num([1, 2])}</c:xVal><c:yVal>${num([3, 4])}</c:yVal>`)}
    <c:axId val="1"/><c:axId val="2"/></c:scatterChart>${valAx(1, 2, 'b')}${valAx(2, 1, 'l')}`);
  const [series] = chart.series;
  assert.deepEqual(series.trendlines, [{ type: 'poly', order: 3, name: '拟合' }, { type: 'movingAvg', period: 4 }]);
  assert.deepEqual(series.errorBars, [
    { direction: 'y', type: 'both', valueType: 'percentage', value: 5 },
    { direction: 'x', type: 'plus', valueType: 'cust', plus: [0.1, 0.2] },
  ]);
});

test('trendline fitting matches the closed forms', () => {
  const near = (actual, expected) => assert.ok(Math.abs(actual - expected) < 1e-6, `${actual} vs ${expected}`);
  // y = 2x + 1 精确拟合；线性趋势线只取两端两个点。
  const linear = trendlinePoints([1, 2, 3, 4].map((x) => ({ x, y: 2 * x + 1 })), 'linear');
  assert.equal(linear.length, 2);
  near(linear[0].y, 3); near(linear[1].y, 9);
  // y = 3·e^(0.5x)
  const exp = trendlinePoints([0, 1, 2, 3].map((x) => ({ x, y: 3 * Math.exp(0.5 * x) })), 'exp', { samples: 2 });
  near(exp[1].y, 3 * Math.exp(1.5));
  // y = x² − 2x + 4，二次多项式
  const poly = trendlinePoints([-2, -1, 0, 1, 2, 3].map((x) => ({ x, y: x * x - 2 * x + 4 })), 'poly', { order: 2, samples: 3 });
  near(poly[1].y, 0.25 - 1 + 4);
  // y = 5·x^1.5
  const power = trendlinePoints([1, 2, 4, 8].map((x) => ({ x, y: 5 * x ** 1.5 })), 'power', { samples: 2 });
  near(power[1].y, 5 * 8 ** 1.5);
  // log：y = 2 + 3 ln x
  const log = trendlinePoints([1, 2, 3, 4].map((x) => ({ x, y: 2 + 3 * Math.log(x) })), 'log', { samples: 2 });
  near(log[0].y, 2);
  // 滑动平均从第 period 个点起才有值。
  assert.deepEqual(trendlinePoints([1, 2, 3, 4].map((x, index) => ({ x, y: [2, 4, 6, 8][index] })), 'movingAvg', { period: 2 }).map((point) => point.y), [3, 5, 7]);
  // exp 拒绝 y ≤ 0 的点；剩下的点不够就不画。
  assert.deepEqual(trendlinePoints([{ x: 1, y: -1 }, { x: 2, y: 0 }, { x: 3, y: 2 }], 'exp'), []);
  // 所有 x 相同：矩阵奇异，不画而不是画出 NaN。
  assert.deepEqual(trendlinePoints([{ x: 1, y: 1 }, { x: 1, y: 2 }], 'linear'), []);
});

test('error bar amounts follow Excel, including stdDev centred on the series mean', () => {
  const values = [2, 4, null, 6];
  assert.deepEqual(errorBarRanges(values, { type: 'both', valueType: 'fixedVal', value: 1 }),
    [{ low: 1, high: 3 }, { low: 3, high: 5 }, null, { low: 5, high: 7 }]);
  assert.deepEqual(errorBarRanges([10], { type: 'plus', valueType: 'percentage', value: 20 }), [{ low: 10, high: 12 }]);
  assert.deepEqual(errorBarRanges([10, 20], { type: 'minus', valueType: 'cust', minus: [1, 3] }), [{ low: 9, high: 10 }, { low: 17, high: 20 }]);
  // 样本标准差：[2,4,6] 是 2；以均值 4 为中心，每个点画的是同一段 [2, 6]。
  const deviation = errorBarRanges([2, 4, 6], { type: 'both', valueType: 'stdDev', value: 1 });
  assert.deepEqual(deviation, [{ low: 2, high: 6 }, { low: 2, high: 6 }, { low: 2, high: 6 }]);
  const standardError = errorBarRanges([2, 4, 6], { type: 'both', valueType: 'stdErr' });
  assert.ok(Math.abs(standardError[0].high - (2 + 2 / Math.sqrt(3))) < 1e-9);
});

test('radar spokes start straight up and go clockwise', () => {
  const top = radarPoint(0, 4, 10, 0, 0);
  const right = radarPoint(1, 4, 10, 0, 0);
  assert.ok(Math.abs(top.x) < 1e-9 && Math.abs(top.y + 10) < 1e-9);
  assert.ok(Math.abs(right.x - 10) < 1e-9 && Math.abs(right.y) < 1e-9);
});
