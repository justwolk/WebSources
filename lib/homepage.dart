// ignore_for_file: prefer_const_constructors, prefer_const_literals_to_create_immutables

import 'package:flutter/material.dart';
import 'dart:convert';
import 'package:dio/dio.dart';

class HomePage extends StatefulWidget {
  const HomePage({super.key});

  @override
  State<HomePage> createState() => _HomePageState();
}

class _HomePageState extends State<HomePage> {
  bool? isChecked = true;
  List<String> _selectedItems = [];
  late List<Song> _songs;
  final dio = Dio();

  @override
  void initState() {
    _fetchSongs();
    super.initState();
  }

  Future<void> _fetchSongs() async {
    try {
      final response =
          await dio.get('http://g90423oo.beget.tech/images/test.json');

      if (response.statusCode == 200) {
        final jsonData = response.data as List<dynamic>;
        setState(() {
          _songs = jsonData
              .map((songJson) => Song.fromJson(songJson['song_data']))
              .toList();
        });
      } else {
        throw Exception('Failed to load songs');
      }
    } catch (e) {
      print(e.toString());
    }
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      backgroundColor: Colors.white,
      drawer: Drawer(),
      body: SingleChildScrollView(
        child: Padding(
          padding: const EdgeInsets.all(20.0),
          child: Container(
            child: Column(children: <Widget>[
              SizedBox(height: 10),
              SelectableText("Оставить заявку на прохождение программ ДПО",
                  style: TextStyle(fontSize: 20)),
              SizedBox(height: 20),
              TextField(
                decoration: InputDecoration(
                  border: OutlineInputBorder(),
                  labelText: 'Полное ФИО',
                ),
              ),
              SizedBox(height: 10),
              TextField(
                decoration: InputDecoration(
                  border: OutlineInputBorder(),
                  labelText: 'Должность',
                ),
              ),
              SizedBox(height: 10),
              TextField(
                decoration: InputDecoration(
                  border: OutlineInputBorder(),
                  labelText: 'Адрес эл. почты',
                ),
              ),
              SizedBox(height: 10),
              TextField(
                decoration: InputDecoration(
                  border: OutlineInputBorder(),
                  labelText: 'Желаемый год обучения',
                ),
              ),
              SizedBox(height: 10),
              TextField(
                decoration: InputDecoration(
                  border: OutlineInputBorder(),
                  labelText: 'Другие пожелания',
                ),
              ),
              SizedBox(height: 10),
              Row(
                children: [
                  SelectableText(
                      "Соглашаюсь с политикой обработки персональных данных"),
                  SizedBox(width: 10),
                  Checkbox(
                      value: isChecked,
                      onChanged: (newBool) {
                        // setState(() {
                        //   isChecked = newBool;
                        // });
                      })
                ],
              ),
              ElevatedButton(
                onPressed: () {},
                child: Text("Отправить", style: TextStyle(fontSize: 20)),
              ),
              SizedBox(height: 10),
              SelectableText("Выберите программы ДПО (не более 3)",
                  style: TextStyle(fontSize: 16)),
              SizedBox(height: 10),
              TextField(
                decoration: InputDecoration(
                  border: UnderlineInputBorder(),
                  labelText: 'Введите для поиска',
                ),
              ),
              SizedBox(height: 10),
              ListView.builder(
                  shrinkWrap: true,
                  itemCount: 10,
                  itemBuilder: (BuildContext context, int index) {
                    return ListTile(
                      selected: false,
                      title: Text("Программа ДПО"),
                      subtitle: Text("Описание программы ДПО"),
                      trailing: Icon(Icons.arrow_forward_ios),
                      onTap: () {
                        setState(() {});
                      },
                    );
                  }),
            ]),
          ),
        ),
      ),
    );
  }
}

class Song {
  final String title;
  final String duration;
  final String additionalInformation;

  Song({
    required this.title,
    required this.duration,
    required this.additionalInformation,
  });

  factory Song.fromJson(Map<String, dynamic> json) {
    return Song(
      title: json['title'],
      duration: json['duration'],
      additionalInformation: json['additional_information'],
    );
  }
}
