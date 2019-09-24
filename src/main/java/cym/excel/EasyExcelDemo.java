package cym.excel;

import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.event.WriteHandler;
import com.alibaba.excel.metadata.BaseRowModel;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.support.ExcelTypeEnum;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.junit.Test;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by Administrator on 2019/9/5.
 */
public class EasyExcelDemo {
    @Test
    public void exportExcel() throws IOException {
        try {
            OutputStream out = new FileOutputStream("withHead12.xlsx");
            ExcelWriter writer = new ExcelWriter(null,out, ExcelTypeEnum.XLSX,true,new MyCellStyle());
            Sheet sheet1 = new Sheet(1, 0, ExcelPropertyIndexModel.class);
            sheet1.setSheetName("sheet1");
            List<ExcelPropertyIndexModel> data = new ArrayList();
            ExcelPropertyIndexModel mode = new ExcelPropertyIndexModel();
            mode.setCity("琼海市");
            mode.setNum(0);
            mode.setTotal(0);
            data.add(mode);
            writer.write(data, sheet1);
            writer.finish();
            out.close();

        } catch (Exception e) {
            e.printStackTrace();
        }

    }
    public static class ExcelPropertyIndexModel extends BaseRowModel {


        @ExcelProperty(value = {"项目","项目","项目","项目","序号"}, index = 0)
        private Integer num=1;

        @ExcelProperty(value = {"代码","代码","代码","代码","市县"}, index = 1)
        private String city;

        @ExcelProperty(value = {"基本医疗保险参保人员人数合计","基本医疗保险参保人员人数合计","基本医疗保险参保人员人数合计","基本医疗保险参保人员人数合计","1"}, index = 2)
        private Integer total=0;

        @ExcelProperty(value = {"职工医疗保险","参保人数小计","参保人数小计","参保人数小计","2"},index = 3)
        private Integer zgTotal=0;

        @ExcelProperty(value = {"职工医疗保险","","农民工人数","农民工人数","3"},index = 4)
        private Integer zgNMGRS=0;

        @ExcelProperty(value = {"职工医疗保险","","灵活就业人员","灵活就业人员","4"},index = 5)
        private Integer zgLHJYRY=0;

        @ExcelProperty(value = {"职工医疗保险","","","实施统账结合人数","5"},index = 6)
        private Integer zgSSTZJHRS=0;

        @ExcelProperty(value = {"职工医疗保险","实施统账结合","职工","职工","6"},index = 8)
        private Integer zgSSTZJHzg=0;

        @ExcelProperty(value = {"职工医疗保险","实施统账结合","退休人员","退休人员","7"},index = 9)
        private Integer zgSSTZJHtxry=0;

        @ExcelProperty(value = {"职工医疗保险","单建统筹基金","职工","职工","8"},index = 10)
        private Integer zgDJTCJJzg=0;

        @ExcelProperty(value = {"职工医疗保险","单建统筹基金","退休人员","退休人员","9"},index = 11)
        private Integer zgDJTCJJtxry=0;

        @ExcelProperty(value = {"职工医疗保险","享受待遇人数","享受待遇人数","享受待遇人数","10"},index = 12)
        private Integer zgXSDYRS=0;

        @ExcelProperty(value = {"城乡居民医疗保险","参保人数小计","参保人数小计","参保人数小计","11"},index = 13)
        private Integer cxTotal=0;

        @ExcelProperty(value = {"城乡居民医疗保险","","农民工人数","农民工人数","12"},index = 14)
        private Integer cxNMGRS=0;

        @ExcelProperty(value = {"城乡居民医疗保险","","灵活就业人数","灵活就业人数","13"},index = 15)
        private Integer cxLHJYRS=0;

        @ExcelProperty(value = {"城乡居民医疗保险","享受待遇人数","享受待遇人数","享受待遇人数","14"},index = 16)
        private Integer cxXSDYRS=0;

        public Integer getNum() {
            return num;
        }

        public void setNum(Integer num) {
            this.num = num;
        }

        public String getCity() {
            return city;
        }

        public void setCity(String city) {
            this.city = city;
        }

        public Integer getTotal() {
            return total;
        }

        public void setTotal(Integer total) {
            this.total = total;
        }

        public Integer getZgTotal() {
            return zgTotal;
        }

        public void setZgTotal(Integer zgTotal) {
            this.zgTotal = zgTotal;
        }

        public Integer getZgNMGRS() {
            return zgNMGRS;
        }

        public void setZgNMGRS(Integer zgNMGRS) {
            this.zgNMGRS = zgNMGRS;
        }

        public Integer getZgLHJYRY() {
            return zgLHJYRY;
        }

        public void setZgLHJYRY(Integer zgLHJYRY) {
            this.zgLHJYRY = zgLHJYRY;
        }

        public Integer getZgSSTZJHRS() {
            return zgSSTZJHRS;
        }

        public void setZgSSTZJHRS(Integer zgSSTZJHRS) {
            this.zgSSTZJHRS = zgSSTZJHRS;
        }

        public Integer getZgSSTZJHzg() {
            return zgSSTZJHzg;
        }

        public void setZgSSTZJHzg(Integer zgSSTZJHzg) {
            this.zgSSTZJHzg = zgSSTZJHzg;
        }

        public Integer getZgSSTZJHtxry() {
            return zgSSTZJHtxry;
        }

        public void setZgSSTZJHtxry(Integer zgSSTZJHtxry) {
            this.zgSSTZJHtxry = zgSSTZJHtxry;
        }

        public Integer getZgDJTCJJzg() {
            return zgDJTCJJzg;
        }

        public void setZgDJTCJJzg(Integer zgDJTCJJzg) {
            this.zgDJTCJJzg = zgDJTCJJzg;
        }

        public Integer getZgDJTCJJtxry() {
            return zgDJTCJJtxry;
        }

        public void setZgDJTCJJtxry(Integer zgDJTCJJtxry) {
            this.zgDJTCJJtxry = zgDJTCJJtxry;
        }

        public Integer getZgXSDYRS() {
            return zgXSDYRS;
        }

        public void setZgXSDYRS(Integer zgXSDYRS) {
            this.zgXSDYRS = zgXSDYRS;
        }

        public Integer getCxTotal() {
            return cxTotal;
        }

        public void setCxTotal(Integer cxTotal) {
            this.cxTotal = cxTotal;
        }

        public Integer getCxNMGRS() {
            return cxNMGRS;
        }

        public void setCxNMGRS(Integer cxNMGRS) {
            this.cxNMGRS = cxNMGRS;
        }

        public Integer getCxLHJYRS() {
            return cxLHJYRS;
        }

        public void setCxLHJYRS(Integer cxLHJYRS) {
            this.cxLHJYRS = cxLHJYRS;
        }

        public Integer getCxXSDYRS() {
            return cxXSDYRS;
        }

        public void setCxXSDYRS(Integer cxXSDYRS) {
            this.cxXSDYRS = cxXSDYRS;
        }
    }

    class MyCellStyle implements WriteHandler {
        @Override
        public void sheet(int i, org.apache.poi.ss.usermodel.Sheet sheet) {

        }

        @Override
        public void row(int i, Row row) {

        }

        @Override
        public void cell(int i, Cell cell) {

        }
    }


}
