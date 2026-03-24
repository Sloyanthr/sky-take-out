package com.sky.service.impl;


import com.sky.dto.GoodsSalesDTO;
import com.sky.entity.Orders;
import com.sky.mapper.OrderMapper;
import com.sky.mapper.UserMapper;
import com.sky.service.ReportService;
import com.sky.service.WorkspaceService;
import com.sky.vo.*;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.InputStream;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

@Service
@Slf4j
public class ReportServiceImpl implements ReportService {

    @Autowired
    private OrderMapper orderMapper;
    @Autowired
    private UserMapper userMapper;
    @Autowired
    private WorkspaceService workspaceService;

    /**
     * 营业额统计
     *
     * @param begin
     * @param end
     * @return
     */
    @Override
    public TurnoverReportVO getTurnoverStatistics(LocalDate begin, LocalDate end) {
        //当前集合用于存放从begin到end每天的日期
        List<LocalDate> dateList = new ArrayList<>();
        dateList.add(begin);
        while (!begin.equals(end)) {
            //日期计算，begin加1天， until()方法返回值是一个LocalDate对象
            begin = begin.plusDays(1);
            dateList.add(begin);
        }
        List<Double> turnoverList = new ArrayList<>();
        //当前集合用于存放从begin到end每天营业额数据
        for (LocalDate date : dateList){
            //查询date日期的营业额数据，具体指当天"已完成"的订单金额合计
            //首先将date转换为带时分秒的
            LocalDateTime beginTime = LocalDateTime.of(date, LocalTime.MIN);
            LocalDateTime endTime = LocalDateTime.of(date, LocalTime.MAX);
            Map<String, Object> map = new HashMap<>();
            map.put("begin", beginTime);
            map.put("end", endTime);
            map.put("status", Orders.COMPLETED);
            Double turnover = orderMapper.sumByMap(map);
            turnover = turnover == null ? 0.0 : turnover;
            turnoverList.add(turnover);
        }

        //构建返回对象
        return TurnoverReportVO.builder()
                .dateList(StringUtils.join(dateList, ","))
                .turnoverList(StringUtils.join(turnoverList, ","))
                .build();
    }

    /**
     * 统计指定时间区间内的用户数据
     *
     * @param begin
     * @param end
     * @return
     */
    @Override
    public UserReportVO getUserStatistics(LocalDate begin, LocalDate end) {
        List<LocalDate> dateList = new ArrayList<>();
        dateList.add(begin);
        while (!begin.equals(end)) {
            //日期计算，begin加1天， until()方法返回值是一个LocalDate对象
            begin = begin.plusDays(1);
            dateList.add(begin);
        }
        //存放每天用户总量
        List<Integer> totalUserList = new ArrayList<>();
        //存放每天新增用户数量
        List<Integer> newUserList = new ArrayList<>();
        for (LocalDate date : dateList){
            //查询date日期的总用户数量
            //首先将date转换为带时分秒的
            LocalDateTime beginTime = LocalDateTime.of(date, LocalTime.MIN);
            LocalDateTime endTime = LocalDateTime.of(date, LocalTime.MAX);
            Map<String, Object> map = new HashMap<>();
            map.put("end", endTime);
            Integer totalUser = userMapper.countByMap(map);

            map.put("begin", beginTime);
            Integer newUser = userMapper.countByMap(map);


            totalUser = totalUser == null ? 0 : totalUser;
            newUser = newUser == null ? 0 : newUser;
            totalUserList.add(totalUser);
            newUserList.add(newUser);
        }
        return UserReportVO.builder()
                .dateList(StringUtils.join(dateList, ","))
                .totalUserList(StringUtils.join(totalUserList, ","))
                .newUserList(StringUtils.join(newUserList, ","))
                .build();

    }

    /**
     * 统计指定时间区间内的订单数据
     *
     * @param begin
     * @param end
     * @return
     */
    @Override
    public OrderReportVO getOrdersStatistics(LocalDate begin, LocalDate end) {
        // 1. 准备日期列表
        List<LocalDate> dateList = new ArrayList<>();
        for (LocalDate d = begin; !d.isAfter(end); d = d.plusDays(1)) {
            dateList.add(d);
        }

        // 2. 准备每日统计列表
        List<Integer> orderCountList = new ArrayList<>(); // 每日订单数
        List<Integer> validOrderCountList = new ArrayList<>(); // 每日有效订单数 (状态为已完成)

        for (LocalDate date : dateList) {
            LocalDateTime beginTime = LocalDateTime.of(date, LocalTime.MIN);
            LocalDateTime endTime = LocalDateTime.of(date, LocalTime.MAX);

            // 查询每日总订单数
            Map<String, Object> map = new HashMap<>();
            map.put("begin", beginTime);
            map.put("end", endTime);
            Integer orderCount = orderMapper.countByMap(map);
            orderCountList.add(orderCount);

            // 查询每日有效订单数 (状态为已完成)
            map.put("status", Orders.COMPLETED);
            Integer validOrderCount = orderMapper.countByMap(map);
            validOrderCountList.add(validOrderCount);
        }

        // 3. 计算时间段内的总计数据
        // 总订单数：对 orderCountList 求和
        Integer totalOrderCount = orderCountList.stream().reduce(Integer::sum).orElse(0);

        // 总有效订单数：对 validOrderCountList 求和
        Integer validOrderCount = validOrderCountList.stream().reduce(Integer::sum).orElse(0);

        // 订单完成率
        double orderCompletionRate = 0.0;
        if (totalOrderCount != 0) {
            orderCompletionRate = validOrderCount.doubleValue() / totalOrderCount;
        }

        // 4. 封装 VO
        return OrderReportVO.builder()
                .dateList(StringUtils.join(dateList, ","))
                .orderCountList(StringUtils.join(orderCountList, ","))
                .validOrderCountList(StringUtils.join(validOrderCountList, ","))
                .totalOrderCount(totalOrderCount)
                .validOrderCount(validOrderCount)
                .orderCompletionRate(orderCompletionRate)
                .build();
    }

    /**
     * 统计指定时间区间内的销量排名top10
     *
     * @param begin
     * @param end
     * @return
     */
    @Override
    public SalesTop10ReportVO getSalesTop10(LocalDate begin, LocalDate end) {
        LocalDateTime beginTime = LocalDateTime.of(begin, LocalTime.MIN);
        LocalDateTime endTime = LocalDateTime.of(end, LocalTime.MAX);

        Map map = new HashMap();
        map.put("begin", beginTime);
        map.put("end", endTime);
        map.put("status", Orders.COMPLETED);

        // 1. 调用Mapper查询前10名数据
        List<GoodsSalesDTO> goodsSalesDTOList = orderMapper.getSalesTop10(map);

        // 2. 将 List<DTO> 转换为前端需要的两个以逗号隔开的字符串
        // 提取名称列表
        List<String> nameList = goodsSalesDTOList.stream()
                .map(GoodsSalesDTO::getName)
                .collect(Collectors.toList());
        String names = StringUtils.join(nameList, ",");

        // 提取销量列表
        List<Integer> numberList = goodsSalesDTOList.stream()
                .map(GoodsSalesDTO::getNumber)
                .collect(Collectors.toList());
        String numbers = StringUtils.join(numberList, ",");

        // 3. 封装VO返回
        return SalesTop10ReportVO.builder()
                .nameList(names)
                .numberList(numbers)
                .build();
    }

    @Override
    public void exportBusinessData(HttpServletResponse response) {
        //查询数据库，获取营业数据
        LocalDate dateBegin = LocalDate.now().minusDays(30);
        LocalDate dateEnd = LocalDate.now().minusDays(1);
        //查询概览数据
        BusinessDataVO businessData = workspaceService.getBusinessData(
                LocalDateTime.of(dateBegin, LocalTime.MIN),
                LocalDateTime.of(dateEnd, LocalTime.MAX));
        //通过 POI 将数据写到Excel文件中
        InputStream in = this.getClass()
                .getClassLoader().getResourceAsStream("template/运营数据报表模板.xlsx");
        try {
            //基于模板创建一个 Excel 文件
            XSSFWorkbook excel = new XSSFWorkbook(in);

            //填充数据：时间
            XSSFSheet sheet = excel.getSheet("Sheet1");
            sheet.getRow(1).getCell(1)
                    .setCellValue("时间：" + dateBegin + "至" + dateEnd);

            //填充数据：概览数据
            //获得第四行
            XSSFRow row = sheet.getRow(3);
            // 1. 营业额 (B4 -> index: Row 3, Cell 1)
            row.getCell(2).setCellValue(businessData.getTurnover());
            // 2. 订单完成率 (E4 -> index: Row 3, Cell 4)
            row.getCell(4).setCellValue(businessData.getOrderCompletionRate());
            // 3. 新增用户数 (H4 -> index: Row 3, Cell 7)
            row.getCell(6).setCellValue(businessData.getNewUsers());

            //获得第五行
            row = sheet.getRow(4);
            // 4. 有效订单 (B5 -> index: Row 4, Cell 1)
            row.getCell(2).setCellValue(businessData.getValidOrderCount());
            // 5. 平均客单价 (E5 -> index: Row 4, Cell 4)
            row.getCell(4).setCellValue(businessData.getUnitPrice());

            //填充数据：明细数据
            for (int i = 0; i < 30; i++){
                LocalDate date = dateBegin.plusDays(i);
                //查询某一天的营业数据
                BusinessDataVO businessDataOneDay = workspaceService.getBusinessData(
                        LocalDateTime.of(date, LocalTime.MIN),
                        LocalDateTime.of(date, LocalTime.MAX));

                //定位到数据对应行
                row = sheet.getRow(i + 7);
                row.createCell(1).setCellValue(date.toString());                            // 日期 (A列)
                row.createCell(2).setCellValue(businessDataOneDay.getTurnover());           // 营业额 (B列)
                row.createCell(3).setCellValue(businessDataOneDay.getValidOrderCount());    // 有效订单 (C列)
                row.createCell(4).setCellValue(businessDataOneDay.getOrderCompletionRate());// 订单完成率 (D列)
                row.createCell(5).setCellValue(businessDataOneDay.getUnitPrice());          // 平均客单价 (E列)
                row.createCell(6).setCellValue(businessDataOneDay.getNewUsers());           // 新增用户数 (F列)
            }

            //通过输出流将 Excel 文件下载到客户端浏览器上
            ServletOutputStream out = response.getOutputStream();
            excel.write(out);

            //关闭流
            out.close();
            excel.close();

        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
