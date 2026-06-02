(function () {
    const palette = ["#0f766e", "#b7791f", "#b42318", "#4f6f52", "#6b5f2a", "#3f6f90", "#8a4b35"];

    function resizeCanvas(canvas) {
        const rect = canvas.getBoundingClientRect();
        const ratio = window.devicePixelRatio || 1;
        canvas.width = Math.max(1, Math.floor(rect.width * ratio));
        canvas.height = Math.max(1, Math.floor(rect.height * ratio));
        const ctx = canvas.getContext("2d");
        ctx.setTransform(ratio, 0, 0, ratio, 0, 0);
        return { ctx, width: rect.width, height: rect.height };
    }

    function drawNoData(ctx, width, height) {
        ctx.fillStyle = "#65716e";
        ctx.font = "14px Segoe UI, Arial";
        ctx.textAlign = "center";
        ctx.fillText("Sem dados", width / 2, height / 2);
    }

    function drawBar(canvas, items) {
        const { ctx, width, height } = resizeCanvas(canvas);
        ctx.clearRect(0, 0, width, height);
        if (!items.length) return drawNoData(ctx, width, height);
        const padding = { top: 18, right: 12, bottom: 52, left: 36 };
        const max = Math.max(...items.map((item) => item.value), 1);
        const chartWidth = width - padding.left - padding.right;
        const chartHeight = height - padding.top - padding.bottom;
        const slot = chartWidth / items.length;
        ctx.strokeStyle = "#d9e1de";
        ctx.beginPath();
        ctx.moveTo(padding.left, padding.top);
        ctx.lineTo(padding.left, padding.top + chartHeight);
        ctx.lineTo(width - padding.right, padding.top + chartHeight);
        ctx.stroke();
        items.forEach((item, index) => {
            const barHeight = (item.value / max) * (chartHeight - 8);
            const x = padding.left + index * slot + slot * 0.18;
            const y = padding.top + chartHeight - barHeight;
            const barWidth = Math.max(14, slot * 0.64);
            ctx.fillStyle = palette[index % palette.length];
            ctx.fillRect(x, y, barWidth, barHeight);
            ctx.fillStyle = "#17211f";
            ctx.font = "12px Segoe UI, Arial";
            ctx.textAlign = "center";
            ctx.fillText(item.value, x + barWidth / 2, y - 5);
            ctx.save();
            ctx.translate(x + barWidth / 2, padding.top + chartHeight + 12);
            ctx.rotate(-Math.PI / 5);
            ctx.fillStyle = "#65716e";
            ctx.fillText(String(item.label).slice(0, 18), 0, 0);
            ctx.restore();
        });
    }

    function drawPie(canvas, items) {
        const { ctx, width, height } = resizeCanvas(canvas);
        ctx.clearRect(0, 0, width, height);
        const total = items.reduce((sum, item) => sum + item.value, 0);
        if (!total) return drawNoData(ctx, width, height);
        const radius = Math.min(width, height) * 0.32;
        const cx = width * 0.36;
        const cy = height * 0.5;
        let angle = -Math.PI / 2;
        items.forEach((item, index) => {
            const slice = (item.value / total) * Math.PI * 2;
            ctx.beginPath();
            ctx.moveTo(cx, cy);
            ctx.arc(cx, cy, radius, angle, angle + slice);
            ctx.closePath();
            ctx.fillStyle = palette[index % palette.length];
            ctx.fill();
            angle += slice;
        });
        ctx.font = "13px Segoe UI, Arial";
        items.forEach((item, index) => {
            const x = width * 0.68;
            const y = height * 0.32 + index * 24;
            ctx.fillStyle = palette[index % palette.length];
            ctx.fillRect(x, y - 10, 12, 12);
            ctx.fillStyle = "#17211f";
            ctx.textAlign = "left";
            ctx.fillText(`${item.label}: ${item.value}`, x + 18, y);
        });
    }

    function drawLine(canvas, items) {
        const { ctx, width, height } = resizeCanvas(canvas);
        ctx.clearRect(0, 0, width, height);
        if (!items.length) return drawNoData(ctx, width, height);
        const padding = { top: 20, right: 20, bottom: 36, left: 36 };
        const max = Math.max(...items.map((item) => item.value), 1);
        const chartWidth = width - padding.left - padding.right;
        const chartHeight = height - padding.top - padding.bottom;
        ctx.strokeStyle = "#d9e1de";
        ctx.beginPath();
        ctx.moveTo(padding.left, padding.top);
        ctx.lineTo(padding.left, padding.top + chartHeight);
        ctx.lineTo(width - padding.right, padding.top + chartHeight);
        ctx.stroke();
        const points = items.map((item, index) => {
            const x = padding.left + (items.length === 1 ? chartWidth / 2 : (index / (items.length - 1)) * chartWidth);
            const y = padding.top + chartHeight - (item.value / max) * (chartHeight - 8);
            return { x, y, item };
        });
        ctx.strokeStyle = "#0f766e";
        ctx.lineWidth = 3;
        ctx.beginPath();
        points.forEach((point, index) => {
            if (index === 0) ctx.moveTo(point.x, point.y);
            else ctx.lineTo(point.x, point.y);
        });
        ctx.stroke();
        points.forEach((point) => {
            ctx.fillStyle = "#ffffff";
            ctx.strokeStyle = "#0f766e";
            ctx.lineWidth = 3;
            ctx.beginPath();
            ctx.arc(point.x, point.y, 5, 0, Math.PI * 2);
            ctx.fill();
            ctx.stroke();
            ctx.fillStyle = "#17211f";
            ctx.font = "12px Segoe UI, Arial";
            ctx.textAlign = "center";
            ctx.fillText(point.item.value, point.x, point.y - 10);
            ctx.fillStyle = "#65716e";
            ctx.fillText(point.item.label, point.x, padding.top + chartHeight + 20);
        });
    }

    function renderCharts() {
        const dataNode = document.getElementById("chart-data");
        if (!dataNode) return;
        let data = {};
        try {
            data = JSON.parse(dataNode.textContent);
        } catch (error) {
            return;
        }
        document.querySelectorAll("canvas[data-chart]").forEach((canvas) => {
            const source = canvas.dataset.source;
            const items = data[source] || [];
            if (canvas.dataset.chart === "bar") drawBar(canvas, items);
            if (canvas.dataset.chart === "pie") drawPie(canvas, items);
            if (canvas.dataset.chart === "line") drawLine(canvas, items);
        });
    }

    document.querySelectorAll("form[data-confirm]").forEach((form) => {
        form.addEventListener("submit", (event) => {
            if (!window.confirm(form.dataset.confirm)) {
                event.preventDefault();
            }
        });
    });

    window.addEventListener("resize", renderCharts);
    renderCharts();
})();
