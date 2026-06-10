import 'package:flutter/material.dart';
import 'package:flutter/services.dart';

/// A button that copies [code] to the clipboard and briefly confirms.
class CopyCodeButton extends StatefulWidget {
  const CopyCodeButton({super.key, required this.code});

  final String code;

  @override
  State<CopyCodeButton> createState() => _CopyCodeButtonState();
}

class _CopyCodeButtonState extends State<CopyCodeButton> {
  bool _copied = false;

  Future<void> _copy() async {
    await Clipboard.setData(ClipboardData(text: widget.code));
    if (!mounted) return;
    setState(() => _copied = true);
    Future.delayed(const Duration(seconds: 2), () {
      if (mounted) setState(() => _copied = false);
    });
  }

  @override
  Widget build(BuildContext context) {
    return OutlinedButton.icon(
      onPressed: _copy,
      icon: Icon(_copied ? Icons.check : Icons.copy_outlined, size: 18),
      label: Text(_copied ? 'Copied!' : 'Copy code'),
      style: OutlinedButton.styleFrom(
        padding: const EdgeInsets.symmetric(horizontal: 18, vertical: 16),
      ),
    );
  }
}
