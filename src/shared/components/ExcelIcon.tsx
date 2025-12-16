/**
 * Excel icon component using the Microsoft Office Excel logo
 */

import excelIconUrl from "../../assets/20px-Microsoft_Office_Excel_(2025–present).svg.png";

interface ExcelIconProps {
	/** Size in pixels (default: 20) */
	size?: number;
	/** Additional CSS class */
	className?: string;
}

/**
 * Excel icon for export buttons
 * Uses the official Microsoft Office Excel 2025 logo
 */
export function ExcelIcon({ size = 20, className }: ExcelIconProps) {
	return (
		<img
			src={excelIconUrl}
			alt=""
			width={size}
			height={size}
			className={className}
			style={{ display: "block" }}
			aria-hidden="true"
		/>
	);
}
