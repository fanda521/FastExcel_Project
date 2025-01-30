package com.jeffrey.easyexcelstudyone.entity;

import io.swagger.v3.oas.annotations.media.Schema;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.io.Serializable;
import java.math.BigDecimal;

/**
 * @author jeffrey
 * @version 1.0
 * @date 2024/6/17
 * @time 17:14
 * @week 星期一
 * @description 仓位的dto
 **/

@Data
@AllArgsConstructor
@NoArgsConstructor
@Builder
public class PosmWmsBinImportEntity implements Serializable {
    private static final long serialVersionUID = 1L;

    @Schema(description = "rcs仓位的编号")
    private String rcsCode;

    @Schema(description = "仓位名称")
    private String binName;

    @Schema(description = "最大长度")
    private BigDecimal maxLength;

    @Schema(description = "最大宽度")
    private BigDecimal maxWidth;

    @Schema(description = "最大高度")
    private BigDecimal maxHeight;

    @Schema(description = "最大重量")
    private BigDecimal maxWeight;

    @Schema(description = "状态")
    private String status;

    @Schema(description = "平面单位")
    private String planeUom;

    @Schema(description = "重量单位")
    private String weightUom;

    @Schema(description = "x")
    private String x;

    @Schema(description = "y")
    private String y;

    @Schema(description = "z")
    private String z;

    @Schema(description = "仓位类型")
    private String binType;

    @Schema(description = "区域编号")
    private String zoneCode;

    @Schema(description = "优先级")
    private String priority;

    @Schema(description = "货架类型")
    private String shelfType;

    @Schema(description = "上架策略区域")
    private String putawayStrategyZone;


}
