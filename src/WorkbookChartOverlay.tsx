import type { DataEditorRef } from "@glideapps/glide-data-grid";
import type { EChartsOption } from "echarts";
import ReactECharts from "echarts-for-react";
import type { PointerEvent, RefObject } from "react";
import { useMemo, useState } from "react";

import {
  MIN_CHART_LAYOUT_HEIGHT,
  MIN_CHART_LAYOUT_WIDTH,
  type WorkbookChartLayout,
  type WorkbookChartPreview,
} from "./workbook-core";

const GRID_COLUMN_WIDTH = 140;
const GRID_ROW_HEIGHT = 34;

type ChartInteraction = {
  chartId: string;
  kind: "move" | "resize";
  originClientX: number;
  originClientY: number;
  originLayout: WorkbookChartLayout;
  pointerId: number;
} | null;

interface WorkbookChartOverlayProps {
  gridRef: RefObject<DataEditorRef>;
  isLoading: boolean;
  onCommitChartLayout: (chartId: string, layout: WorkbookChartLayout) => Promise<void>;
  onEditChart: (chartId: string) => void;
  onSelectChart: (chartId: string | null) => void;
  previews: WorkbookChartPreview[];
  selectedChartId: string | null;
  surfaceRef: RefObject<HTMLElement>;
  viewportNonce: number;
}

export function WorkbookChartOverlay({
  gridRef,
  isLoading,
  onCommitChartLayout,
  onEditChart,
  onSelectChart,
  previews,
  selectedChartId,
  surfaceRef,
}: WorkbookChartOverlayProps) {
  const [interaction, setInteraction] = useState<ChartInteraction>(null);
  const [previewLayout, setPreviewLayout] = useState<WorkbookChartLayout | null>(null);
  const sortedPreviews = useMemo(
    () => [...previews].sort((left, right) => left.chart.layout.zIndex - right.chart.layout.zIndex),
    [previews],
  );

  if (previews.length === 0 && !isLoading) {
    return null;
  }

  const getRenderLayout = (preview: WorkbookChartPreview) =>
    interaction?.chartId === preview.chart.id && previewLayout
      ? previewLayout
      : preview.chart.layout;

  const getOverlayRect = (layout: WorkbookChartLayout) => {
    const cellBounds = gridRef.current?.getBounds(layout.startColumn, layout.startRow);
    const surfaceBounds = surfaceRef.current?.getBoundingClientRect();

    if (!cellBounds || !surfaceBounds) {
      return null;
    }

    return {
      height: layout.height,
      left: cellBounds.x - surfaceBounds.left + layout.offsetX,
      top: cellBounds.y - surfaceBounds.top + layout.offsetY,
      width: layout.width,
    };
  };

  const startInteraction = (
    event: PointerEvent<HTMLElement>,
    chartId: string,
    kind: "move" | "resize",
    layout: WorkbookChartLayout,
  ) => {
    event.preventDefault();
    event.stopPropagation();
    event.currentTarget.setPointerCapture(event.pointerId);
    onSelectChart(chartId);
    setInteraction({
      chartId,
      kind,
      originClientX: event.clientX,
      originClientY: event.clientY,
      originLayout: layout,
      pointerId: event.pointerId,
    });
    setPreviewLayout(layout);
  };

  const updateInteraction = (event: PointerEvent<HTMLElement>) => {
    if (!interaction || interaction.pointerId !== event.pointerId) {
      return;
    }

    event.preventDefault();
    event.stopPropagation();
    setPreviewLayout(calculateInteractionLayout(interaction, event));
  };

  const finishInteraction = (event: PointerEvent<HTMLElement>) => {
    if (!interaction || interaction.pointerId !== event.pointerId) {
      return;
    }

    event.preventDefault();
    event.stopPropagation();

    const nextLayout = calculateInteractionLayout(interaction, event);

    event.currentTarget.releasePointerCapture(event.pointerId);
    setInteraction(null);
    setPreviewLayout(null);

    void onCommitChartLayout(interaction.chartId, nextLayout);
  };

  return (
    <div className="chart-overlay-layer" aria-label="Embedded charts">
      {isLoading ? <div className="chart-overlay__loading">Syncing</div> : null}
      {sortedPreviews.map((preview) => {
        const layout = getRenderLayout(preview);
        const rect = getOverlayRect(layout);

        if (!rect) {
          return null;
        }

        const isSelected = preview.chart.id === selectedChartId;
        const option = preview.option as EChartsOption | undefined;

        return (
          <article
            aria-label={preview.chart.name}
            className={`chart-overlay-card${isSelected ? " is-selected" : ""}`}
            key={preview.chart.id}
            onDoubleClick={() => {
              onEditChart(preview.chart.id);
            }}
            onPointerDown={(event) => {
              event.stopPropagation();
              onSelectChart(preview.chart.id);
            }}
            onPointerMove={updateInteraction}
            onPointerUp={finishInteraction}
            style={{
              height: rect.height,
              left: rect.left,
              top: rect.top,
              width: rect.width,
              zIndex: isSelected ? 100000 + layout.zIndex : Math.max(1, layout.zIndex + 1),
            }}
          >
            <header
              className="chart-overlay-card__header"
              onPointerDown={(event) => {
                startInteraction(event, preview.chart.id, "move", layout);
              }}
            >
              <strong>{preview.chart.name}</strong>
              <button
                className="chart-overlay-card__button"
                onClick={(event) => {
                  event.stopPropagation();
                  onEditChart(preview.chart.id);
                }}
                onDoubleClick={(event) => {
                  event.stopPropagation();
                }}
                onPointerDown={(event) => {
                  event.stopPropagation();
                  onSelectChart(preview.chart.id);
                }}
                type="button"
              >
                Edit
              </button>
            </header>
            <div className="chart-overlay-card__body">
              {preview.status === "ok" && option ? (
                <ReactECharts
                  className="chart-overlay-card__echarts"
                  lazyUpdate
                  notMerge
                  option={option}
                  style={{ height: "100%", width: "100%" }}
                />
              ) : (
                <div className="chart-overlay-card__invalid">Invalid chart</div>
              )}
            </div>
            <button
              aria-label={`Resize ${preview.chart.name}`}
              className="chart-overlay-card__resize"
              onPointerDown={(event) => {
                startInteraction(event, preview.chart.id, "resize", layout);
              }}
              type="button"
            />
          </article>
        );
      })}
    </div>
  );
}

function calculateInteractionLayout(
  interaction: Exclude<ChartInteraction, null>,
  event: PointerEvent<HTMLElement>,
): WorkbookChartLayout {
  const deltaX = event.clientX - interaction.originClientX;
  const deltaY = event.clientY - interaction.originClientY;

  return interaction.kind === "move"
    ? normalizeMovedLayout(interaction.originLayout, deltaX, deltaY)
    : {
        ...interaction.originLayout,
        height: Math.max(
          MIN_CHART_LAYOUT_HEIGHT,
          Math.round(interaction.originLayout.height + deltaY),
        ),
        width: Math.max(
          MIN_CHART_LAYOUT_WIDTH,
          Math.round(interaction.originLayout.width + deltaX),
        ),
      };
}

function normalizeMovedLayout(
  layout: WorkbookChartLayout,
  deltaX: number,
  deltaY: number,
): WorkbookChartLayout {
  const horizontal = normalizeAxisPosition(
    layout.startColumn,
    layout.offsetX + deltaX,
    GRID_COLUMN_WIDTH,
  );
  const vertical = normalizeAxisPosition(layout.startRow, layout.offsetY + deltaY, GRID_ROW_HEIGHT);

  return {
    ...layout,
    offsetX: horizontal.offset,
    offsetY: vertical.offset,
    startColumn: horizontal.index,
    startRow: vertical.index,
  };
}

function normalizeAxisPosition(originIndex: number, offset: number, unitSize: number) {
  const indexShift = Math.floor(offset / unitSize);
  const nextIndex = originIndex + indexShift;

  if (nextIndex < 0) {
    return {
      index: 0,
      offset: 0,
    };
  }

  return {
    index: nextIndex,
    offset: Math.round(offset - indexShift * unitSize),
  };
}
