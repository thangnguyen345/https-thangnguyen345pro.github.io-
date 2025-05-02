// Mảng chứa thông tin khách hàng
let customers = [];

// Hàm thêm khách hàng vào mảng
function addCustomer() {
    const customerName = document.getElementById('customerName').value;
    const customerPhone = document.getElementById('customerPhone').value;
    const customerAddress = document.getElementById('customerAddress').value;

    if (customerName && customerPhone && customerAddress) {
        customers.push({ name: customerName, phone: customerPhone, address: customerAddress });

        // Sau khi thêm khách hàng, xóa dữ liệu trong ô nhập
        document.getElementById('customerName').value = '';
        document.getElementById('customerPhone').value = '';
        document.getElementById('customerAddress').value = '';
        alert('Khách hàng đã được thêm!');
    } else {
        alert('Vui lòng nhập đầy đủ thông tin khách hàng!');
    }
}

// Mảng chứa sản phẩm
let products = [];

// Hàm thêm sản phẩm vào bảng
function addProduct() {
    const productName = document.getElementById('productName').value;
    const productPrice = parseFloat(document.getElementById('productPrice').value);
    const productQuantity = parseInt(document.getElementById('productQuantity').value);

    if (productName && !isNaN(productPrice) && !isNaN(productQuantity) && productQuantity > 0) {
        const totalPrice = productPrice * productQuantity;

        products.push({ name: productName, price: productPrice, quantity: productQuantity, total: totalPrice });

        const table = document.getElementById('productTable');
        const row = table.insertRow();
        row.innerHTML = `
            <td>${productName}</td>
            <td>${productPrice.toFixed(2)}</td>
            <td>${productQuantity}</td>
            <td>${totalPrice.toFixed(2)}</td>
            <td><button class="button-delete" onclick="deleteProduct(this)">Xóa</button></td>
        `;

        document.getElementById('productName').value = '';
        document.getElementById('productPrice').value = '';
        document.getElementById('productQuantity').value = '';

        calculateTotal();
    } else {
        alert("Vui lòng nhập đầy đủ thông tin sản phẩm.");
    }
}

// Hàm tính tổng tiền
function calculateTotal() {
    const rows = document.getElementById('productTable').rows;
    let total = 0;

    for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        const totalPrice = parseFloat(row.cells[3].textContent);
        total += totalPrice;
    }

    document.getElementById('finalTotal').textContent = total.toFixed(2);
}

// Hàm xóa sản phẩm
function deleteProduct(button) {
    const row = button.closest('tr');
    const rowIndex = row.rowIndex - 1;  // Trừ đi 1 vì bảng có header
    products.splice(rowIndex, 1);
    row.remove();
    calculateTotal();
}

// Hàm lưu tất cả thông tin vào Excel
function saveAllToExcel() {
    const wb = XLSX.utils.book_new();

    // Lưu thông tin khách hàng
    if (customers.length > 0) {
        const customerData = [
            ["Tên khách hàng", "Số điện thoại", "Địa chỉ"]
        ];

        customers.forEach(customer => {
            customerData.push([customer.name, customer.phone, customer.address]);
        });

        const wsCustomers = XLSX.utils.aoa_to_sheet(customerData);
        XLSX.utils.book_append_sheet(wb, wsCustomers, "Thông Tin Khách Hàng");
    }

    // Lưu danh sách sản phẩm
    if (products.length > 0) {
        const productData = [
            ["Tên Sản Phẩm", "Giá", "Số Lượng", "Tổng Tiền"]
        ];

        products.forEach(product => {
            productData.push([product.name, product.price, product.quantity, product.total]);
        });

        const wsProducts = XLSX.utils.aoa_to_sheet(productData);
        XLSX.utils.book_append_sheet(wb, wsProducts, "Danh Sách Sản Phẩm");
    }

    // Xuất file Excel
    XLSX.writeFile(wb, "ThongTinKhachHangVaSanPham.xlsx");
}
