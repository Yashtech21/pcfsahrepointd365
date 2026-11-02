import * as React from "react";

/**
 * Width of an element in px, tracked with ResizeObserver (the control's own element only; no page DOM access).
 * Updates are batched to one per animation frame. Returns undefined until first measured.
 */
export function useElementWidth<T extends HTMLElement>(): [React.RefObject<T>, number | undefined] {
    const ref = React.useRef<T>(null);
    const [width, setWidth] = React.useState<number>();
    React.useEffect(() => {
        const el = ref.current;
        if (!el || typeof ResizeObserver === "undefined") return;
        let frame = 0;
        const ro = new ResizeObserver(entries => {
            const w = Math.round(entries[0].contentRect.width);
            cancelAnimationFrame(frame);
            frame = requestAnimationFrame(() => setWidth(prev => (prev === w ? prev : w)));
        });
        ro.observe(el);
        return () => { cancelAnimationFrame(frame); ro.disconnect(); };
    }, []);
    return [ref, width];
}
