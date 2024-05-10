import 'package:flutter/material.dart';

class CustomIconButton extends StatelessWidget {
  final String text;
  final VoidCallback onPressed;
  final IconData iconData;
  final Gradient gradient; // Add gradient property
  final Color textColor;

  const CustomIconButton({super.key,
    required this.text,
    required this.onPressed,
    required this.iconData,
    this.gradient = const LinearGradient(
      begin: Alignment.topLeft,
      end: Alignment.bottomRight,
      colors: [
        Color.fromARGB(255, 102, 0, 166),
        Color.fromARGB(255, 247, 27, 199)
      ],
    ), // Default gradient
    this.textColor = Colors.white,
  });

  @override
  Widget build(BuildContext context) {
    return Padding(
      padding: const EdgeInsets.symmetric(horizontal: 26.0),
      child: DecoratedBox(
        decoration: BoxDecoration(
          gradient: gradient, // Apply gradient here
          borderRadius: BorderRadius.circular(100.0),
        ),
        child: ElevatedButton(
          onPressed: onPressed,
          style: ElevatedButton.styleFrom(
            padding: const EdgeInsets.symmetric(vertical: 35.0),
            shape: RoundedRectangleBorder(
              borderRadius: BorderRadius.circular(20.0),
            ),
            elevation: 0, // Set elevation to 0 to avoid shadow
            backgroundColor:
            Colors.transparent, // Set background color to transparent
          ),
          child: Row(
            children: [
              Padding(
                padding: const EdgeInsets.only(left: 40.0),
                child: Text(
                  text,
                  style: TextStyle(
                    color: textColor,
                    fontSize: 30.0,
                  ),
                ),
              ),
              Expanded(
                child:
                Container(), // This container expands to push the icon to the right
              ),
              Padding(
                padding: const EdgeInsets.only(right: 16.0),
                child: Icon(
                  iconData,
                  color: textColor,
                  size: 50,
                ),
              ),
            ],
          ),
        ),
      ),
    );
  }
}
