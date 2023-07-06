package DrawLuckyApp;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Random;

public class drawlucky {
    public static ArrayList<String> readExcel(String fileName) {
        ArrayList<String> list = new ArrayList<>(); // Tạo một danh sách để lưu dữ liệu
        try {
            FileInputStream excelFile = new FileInputStream(fileName); // Mở file Excel
            XSSFWorkbook workbook = new XSSFWorkbook(excelFile); // Tạo một đối tượng Workbook
            Sheet sheet = workbook.getSheetAt(0); // Lấy sheet đầu tiên trong Workbook
            for (int i = 0; i <= sheet.getLastRowNum(); i++) { // Duyệt qua các hàng trong sheet
                Row row = sheet.getRow(i); // Lấy hàng theo chỉ số i
                if (row != null) { // Kiểm tra nếu hàng không rỗng
                    String id = ""; // Khởi tạo biến id rỗng
                    String name = ""; // Khởi tạo biến name rỗng
                    String department = ""; // Khởi tạo biến department rỗng
                    Cell cellId = row.getCell(0); // Lấy ô đầu tiên trong hàng (id)
                    if (cellId != null) { // Kiểm tra nếu ô không rỗng
                        if (cellId.getCellType() == CellType.NUMERIC) { // Kiểm tra nếu ô có kiểu số
                            id = String.valueOf(cellId.getNumericCellValue()); // Lấy giá trị số của ô và chuyển thành chuỗi
                        } else {
                            // Xử lý các trường hợp khác
                        }
                    }
                    Cell cellName = row.getCell(1); // Lấy ô thứ hai trong hàng (name)
                    if (cellName != null) { // Kiểm tra nếu ô không rỗng
                        if (cellName.getCellType() == CellType.STRING) { // Kiểm tra nếu ô có kiểu chuỗi
                            name = cellName.getStringCellValue(); // Lấy giá trị chuỗi của ô
                        } else {
                            // Xử lý các trường hợp khác
                        }
                    }
                    Cell cellDepartment = row.getCell(2); // Lấy ô thứ ba trong hàng (department)
                    if (cellDepartment != null) { // Kiểm tra nếu ô không rỗng
                        if (cellDepartment.getCellType() == CellType.STRING) { // Kiểm tra nếu ô có kiểu chuỗi
                            department = cellDepartment.getStringCellValue(); // Lấy giá trị chuỗi của ô
                        } else {
                            // Xử lý các trường hợp khác
                        }
                    }
                    if (!id.isEmpty() && !name.isEmpty() && !department.isEmpty()) { // Kiểm tra nếu cả ba giá trị đều không rỗng
                        list.add(id + " - " + name + " - " + department); // Thêm chuỗi chứa id, name và department vào danh sách
                    }
                }
            }
            workbook.close(); // Đóng Workbook
            excelFile.close(); // Đóng file Excel
        } catch (IOException e) { // Bắt ngoại lệ IOException nếu có lỗi khi đọc hoặc ghi file
            e.printStackTrace(); // In ra thông báo lỗi
        }
        return list; // Trả về danh sách
    }

    public static String drawLucky(ArrayList<String> list) {
        if (list.isEmpty()) { // Kiểm tra nếu danh sách rỗng
            return "Không có dữ liệu"; // Trả về thông báo không có dữ liệu
        }
        Random random = new Random(); // Tạo một đối tượng Random để sinh số ngẫu nhiên
        ArrayList<Integer> indexes = new ArrayList<>(); // Tạo một danh sách để lưu các chỉ số ngẫu nhiên
        String result = ""; // Tạo một chuỗi để lưu kết quả
        int firstPrize = 1; // Số lượng giải nhất
        int secondPrize = 2; // Số lượng giải nhì
        int thirdPrize = 3; // Số lượng giải ba
        int consolationPrize = 5; // Số lượng giải khuyến khích
        int totalPrize = firstPrize + secondPrize + thirdPrize + consolationPrize; // Tổng số lượng giải thưởng
        if (totalPrize > list.size()) { // Kiểm tra nếu tổng số lượng giải thưởng lớn hơn kích thước của danh sách
            return "Số lượng giải thưởng vượt quá số lượng người tham gia"; // Trả về thông báo lỗi
        }
        result += "Giải nhất:\n"; // Thêm tiêu đề giải nhất vào kết quả
        for (int i = 0; i < firstPrize; i++) { // Lặp qua số lượng giải nhất
            int index = random.nextInt(list.size()); // Sinh một số ngẫu nhiên từ 0 đến kích thước của danh sách - 1
            while (indexes.contains(index)) { // Kiểm tra nếu số ngẫu nhiên đã được chọn trước đó
                index = random.nextInt(list.size()); // Sinh lại số ngẫu nhiên mới
            }
            indexes.add(index); // Thêm số ngẫu nhiên vào danh sách các chỉ số đã chọn
            String value = list.get(index); // Lấy phần tử có chỉ số là số ngẫu nhiên trong danh sách
            result += value + "\n"; // Thêm phần tử vào kết quả
        }
        result += "Giải nhì:\n"; // Thêm tiêu đề giải nhì vào kết quả
        for (int i = 0; i < secondPrize; i++) { // Lặp qua số lượng giải nhì
            int index = random.nextInt(list.size()); // Sinh một số ngẫu nhiên từ 0 đến kích thước của danh sách - 1
            while (indexes.contains(index)) { // Kiểm tra nếu số ngẫu nhiên đã được chọn trước đó
                index = random.nextInt(list.size()); // Sinh lại số ngẫu nhiên mới
            }
            indexes.add(index); // Thêm số ngẫu nhiên vào danh sách các chỉ số đã chọn
            String value = list.get(index); // Lấy phần tử có chỉ số là số ngẫu nhiên trong danh sách
            result += value + "\n"; // Thêm phần tử vào kết quả
        }
        result += "Giải ba:\n"; // Thêm tiêu đề giải ba vào kết quả
        for (int i = 0; i < thirdPrize; i++) { // Lặp qua số lượng giải ba
            int index = random.nextInt(list.size()); // Sinh một số ngẫu nhiên từ 0 đến kích thước của danh sách - 1
            while (indexes.contains(index)) { // Kiểm tra nếu số ngẫu nhiên đã được chọn trước đó
                index = random.nextInt(list.size()); // Sinh lại số ngẫu nhiên mới
            }
            indexes.add(index); // Thêm số ngẫu nhiên vào danh sách các chỉ số đã chọn
            String value = list.get(index); // Lấy phần tử có chỉ số là số ngẫu nhiên trong danh sách
            result += value + "\n"; // Thêm phần tử vào kết quả
        }
        result += "Giải khuyến khích:\n"; // Thêm tiêu đề giải khuyến khích vào kết quả
        for (int i = 0; i < consolationPrize; i++) { // Lặp qua số lượng giải khuyến khích
            int index = random.nextInt(list.size()); // Sinh một số ngẫu nhiên từ 0 đến kích thước của danh sách - 1
            while (indexes.contains(index)) { // Kiểm tra nếu số ngẫu nhiên đã được chọn trước đó
                index = random.nextInt(list.size()); // Sinh lại số ngẫu nhiên mới
            }
            indexes.add(index); // Thêm số ngẫu nhiên vào danh sách các chỉ số đã chọn
            String value = list.get(index); // Lấy phần tử có chỉ số là số ngẫu nhiên trong danh sách
            result += value + "\n"; // Thêm phần tử vào kết quả
        }
        return result; // Trả về kết quả
    }


    public static void main(String[] args) {
        String fileName = "C:\\Users\\ASUS\\Documents\\user.xlsx"; // Đường dẫn của file Excel chứa dữ liệu
        ArrayList<String> list = readExcel(fileName); // Đọc dữ liệu từ file Excel và lưu vào danh sách
        String result = drawLucky(list); // Chọn ngẫu nhiên một phần tử trong danh sách
        System.out.println("Kết quả quay số may mắn là: \n" + result); // In ra kết quả quay số may mắn
    }

}
