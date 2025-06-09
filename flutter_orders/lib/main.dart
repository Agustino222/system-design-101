import 'dart:io';
import 'package:excel/excel.dart';
import 'package:flutter/material.dart';

const headers = [
  'ID',
  'Vendedor',
  'Cliente',
  'Material',
  'Cantidad',
  'TipoMadera',
  'Encargado',
  'Especificaciones',
  'FechaCreacion',
  'FechaListo',
  'FechaEntregado'
];

class Order {
  int id;
  String vendedor;
  String cliente;
  String material;
  String cantidad;
  String tipo;
  String encargado;
  String especificaciones;
  DateTime fechaCreacion;
  DateTime? fechaListo;
  DateTime? fechaEntregado;

  Order({
    required this.id,
    required this.vendedor,
    required this.cliente,
    required this.material,
    required this.cantidad,
    required this.tipo,
    required this.encargado,
    required this.especificaciones,
    required this.fechaCreacion,
    this.fechaListo,
    this.fechaEntregado,
  });

  List<dynamic> toRow() {
    return [
      id,
      vendedor,
      cliente,
      material,
      cantidad,
      tipo,
      encargado,
      especificaciones,
      _fmt(fechaCreacion),
      fechaListo != null ? _fmt(fechaListo!) : '',
      fechaEntregado != null ? _fmt(fechaEntregado!) : '',
    ];
  }

  static Order fromRow(List<Data?> row) {
    return Order(
      id: row[0]!.value,
      vendedor: row[1]!.value,
      cliente: row[2]!.value,
      material: row[3]!.value,
      cantidad: row[4]!.value.toString(),
      tipo: row[5]!.value,
      encargado: row[6]!.value,
      especificaciones: row[7]!.value ?? '',
      fechaCreacion: DateTime.parse(row[8]!.value),
      fechaListo:
          row[9] != null && row[9]!.value != '' ? DateTime.parse(row[9]!.value) : null,
      fechaEntregado: row[10] != null && row[10]!.value != ''
          ? DateTime.parse(row[10]!.value)
          : null,
    );
  }
}

String _fmt(DateTime dt) => dt.toIso8601String();

class OrderRepository {
  final String path;
  OrderRepository(this.path);

  List<Order> load() {
    final file = File(path);
    if (!file.existsSync()) {
      final excel = Excel.createExcel();
      excel.updateCell('Sheet1', CellIndex.indexByString('A1'), '');
      excel.encode().then((data) => file.writeAsBytesSync(data));
    }
    final bytes = file.readAsBytesSync();
    final excel = Excel.decodeBytes(bytes);
    final sheet = excel.tables[excel.tables.keys.first]!;
    final rows = sheet.rows.skip(1); // skip header
    final orders = <Order>[];
    for (final r in rows) {
      if (r.isEmpty || r[0] == null) continue;
      orders.add(Order.fromRow(r));
    }
    return orders;
  }

  void save(List<Order> orders) {
    final excel = Excel.createExcel();
    final sheet = excel['Pedidos'];
    sheet.appendRow(headers);
    for (final o in orders) {
      sheet.appendRow(o.toRow());
    }
    final file = File(path);
    file.writeAsBytesSync(excel.encode()!);
  }
}

void main() {
  runApp(OrderApp());
}

class OrderApp extends StatefulWidget {
  @override
  _OrderAppState createState() => _OrderAppState();
}

class _OrderAppState extends State<OrderApp> {
  late OrderRepository repo;
  List<Order> orders = [];
  final String configPath =
      Platform.environment['FERRETERIA_CONFIG_PATH'] ?? 'configuracion.xlsx';

  @override
  void initState() {
    super.initState();
    repo = OrderRepository(configPath);
    orders = repo.load();
  }

  void _addOrder(Order order) {
    setState(() {
      orders.add(order);
      repo.save(orders);
    });
  }

  void _updateOrders() {
    setState(() {
      repo.save(orders);
    });
  }

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      home: Scaffold(
        appBar: AppBar(title: Text('Gestor de Pedidos')),
        body: OrderTabs(
          orders: orders,
          onUpdate: _updateOrders,
        ),
        floatingActionButton: FloatingActionButton(
          onPressed: () async {
            final order = await Navigator.of(context).push<Order>(
              MaterialPageRoute(builder: (_) => OrderForm(nextId: orders.length + 1)),
            );
            if (order != null) _addOrder(order);
          },
          child: Icon(Icons.add),
        ),
      ),
    );
  }
}

class OrderTabs extends StatefulWidget {
  final List<Order> orders;
  final VoidCallback onUpdate;
  OrderTabs({required this.orders, required this.onUpdate});
  @override
  _OrderTabsState createState() => _OrderTabsState();
}

class _OrderTabsState extends State<OrderTabs> {
  @override
  Widget build(BuildContext context) {
    return DefaultTabController(
      length: 3,
      child: Column(
        children: [
          TabBar(tabs: [
            Tab(text: 'Pendientes'),
            Tab(text: 'Listos'),
            Tab(text: 'Entregados'),
          ]),
          Expanded(
            child: TabBarView(
              children: [
                _buildList(context, (o) => o.fechaListo == null),
                _buildList(context, (o) => o.fechaListo != null && o.fechaEntregado == null),
                _buildList(context, (o) => o.fechaEntregado != null),
              ],
            ),
          ),
        ],
      ),
    );
  }

  Widget _buildList(BuildContext context, bool Function(Order) filter) {
    final items = widget.orders.where(filter).toList();
    return ListView.builder(
      itemCount: items.length,
      itemBuilder: (context, idx) {
        final o = items[idx];
        return ListTile(
          title: Text('${o.cliente} (${o.material})'),
          subtitle: Text('ID ${o.id} - ${_fmt(o.fechaCreacion)}'),
          trailing: Row(
            mainAxisSize: MainAxisSize.min,
            children: _buildActions(o),
          ),
        );
      },
    );
  }

  List<Widget> _buildActions(Order o) {
    final List<Widget> actions = [];
    if (o.fechaListo == null) {
      actions.add(IconButton(
          icon: Icon(Icons.check),
          onPressed: () {
            setState(() {
              o.fechaListo = DateTime.now();
            });
            widget.onUpdate();
          }));
    } else if (o.fechaEntregado == null) {
      actions.add(IconButton(
          icon: Icon(Icons.local_shipping),
          onPressed: () {
            setState(() {
              o.fechaEntregado = DateTime.now();
            });
            widget.onUpdate();
          }));
    }
    return actions;
  }
}

class OrderForm extends StatefulWidget {
  final int nextId;
  OrderForm({required this.nextId});
  @override
  _OrderFormState createState() => _OrderFormState();
}

class _OrderFormState extends State<OrderForm> {
  final _formKey = GlobalKey<FormState>();
  final _clienteController = TextEditingController();
  final _materialController = TextEditingController();
  final _cantidadController = TextEditingController();
  final _tipoController = TextEditingController();
  final _encargadoController = TextEditingController();
  final _especificacionesController = TextEditingController();
  final _vendedorController = TextEditingController();

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(title: Text('Nuevo Pedido')),
      body: Padding(
        padding: EdgeInsets.all(16),
        child: Form(
          key: _formKey,
          child: ListView(
            children: [
              TextFormField(
                controller: _vendedorController,
                decoration: InputDecoration(labelText: 'Vendedor'),
              ),
              TextFormField(
                controller: _clienteController,
                decoration: InputDecoration(labelText: 'Cliente'),
                validator: (v) => v == null || v.isEmpty ? 'Requerido' : null,
              ),
              TextFormField(
                controller: _materialController,
                decoration: InputDecoration(labelText: 'Material'),
              ),
              TextFormField(
                controller: _cantidadController,
                decoration: InputDecoration(labelText: 'Cantidad'),
              ),
              TextFormField(
                controller: _tipoController,
                decoration: InputDecoration(labelText: 'Tipo de madera'),
              ),
              TextFormField(
                controller: _encargadoController,
                decoration: InputDecoration(labelText: 'Encargado'),
              ),
              TextFormField(
                controller: _especificacionesController,
                decoration: InputDecoration(labelText: 'Especificaciones'),
              ),
              SizedBox(height: 20),
              ElevatedButton(
                onPressed: () {
                  if (!_formKey.currentState!.validate()) return;
                  final order = Order(
                    id: widget.nextId,
                    vendedor: _vendedorController.text,
                    cliente: _clienteController.text,
                    material: _materialController.text,
                    cantidad: _cantidadController.text,
                    tipo: _tipoController.text,
                    encargado: _encargadoController.text,
                    especificaciones: _especificacionesController.text,
                    fechaCreacion: DateTime.now(),
                  );
                  Navigator.of(context).pop(order);
                },
                child: Text('Guardar'),
              )
            ],
          ),
        ),
      ),
    );
  }
}

