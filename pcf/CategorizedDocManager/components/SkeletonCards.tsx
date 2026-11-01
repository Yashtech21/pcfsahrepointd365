import * as React from "react";
import { Card, Skeleton, SkeletonItem, makeStyles, tokens } from "@fluentui/react-components";

const useStyles = makeStyles({
    card: { minWidth: "220px" },
    head: { display: "flex", gap: tokens.spacingHorizontalS },
    body: { flex: 1, display: "flex", flexDirection: "column", gap: "6px" },
    actions: { display: "flex", gap: tokens.spacingHorizontalM, marginTop: tokens.spacingVerticalS },
});

/** Placeholder cards with the same footprint as DocumentCard, so the grid doesn't jump when data lands. */
export const SkeletonCards: React.FC<{ count?: number; gridClassName: string }> = ({ count = 6, gridClassName }) => {
    const s = useStyles();
    return (
        <div className={gridClassName} aria-busy="true" aria-label="Loading documents">
            {Array.from({ length: count }, (_, i) => (
                <Card key={i} className={s.card} size="small">
                    <Skeleton>
                        <div className={s.head}>
                            <SkeletonItem shape="square" size={32} />
                            <div className={s.body}>
                                <SkeletonItem size={16} style={{ width: "80%" }} />
                                <SkeletonItem size={16} style={{ width: "40%" }} />
                                <SkeletonItem size={12} style={{ width: "70%" }} />
                                <SkeletonItem size={12} style={{ width: "60%" }} />
                            </div>
                        </div>
                        <div className={s.actions}>
                            <SkeletonItem size={16} style={{ width: "60px" }} />
                            <SkeletonItem size={16} style={{ width: "70px" }} />
                            <SkeletonItem size={16} style={{ width: "55px" }} />
                        </div>
                    </Skeleton>
                </Card>
            ))}
        </div>
    );
};

/** Page-shaped placeholder for the preview pane while the embed URL is fetched. */
export const SkeletonPreview: React.FC = () => (
    <Skeleton aria-label="Loading preview" style={{ padding: 16, display: "flex", flexDirection: "column", gap: 10 }}>
        <SkeletonItem size={24} style={{ width: "50%" }} />
        {Array.from({ length: 12 }, (_, i) => <SkeletonItem key={i} size={12} style={{ width: `${70 + ((i * 17) % 30)}%` }} />)}
        <SkeletonItem shape="rectangle" size={128} style={{ width: "100%" }} />
    </Skeleton>
);
