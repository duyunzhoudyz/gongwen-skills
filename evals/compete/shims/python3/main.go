// python3 命令的最小替身：Windows 上 python3 常指向应用商店的占位程序，而部分竞品 skill 的文档
// 一律写 `python3 脚本`（含 `uv run --with python-docx python3 ...`）。本程序在 PATH 中找 python
// （uv run 时就是 uv 临时环境里的 python），原样转发参数、标准输入输出和退出码。
// 构建：go build -o python3.exe main.go
package main

import (
	"fmt"
	"os"
	"os/exec"
)

func main() {
	args := os.Args[1:]
	target, err := exec.LookPath("python")
	if err != nil {
		// 退而求其次：py 启动器，指定 3.x
		if py, e2 := exec.LookPath("py"); e2 == nil {
			target = py
			args = append([]string{"-3"}, args...)
		} else {
			fmt.Fprintln(os.Stderr, "python3 替身：PATH 中找不到 python 或 py")
			os.Exit(127)
		}
	}
	cmd := exec.Command(target, args...)
	cmd.Stdin, cmd.Stdout, cmd.Stderr = os.Stdin, os.Stdout, os.Stderr
	if err := cmd.Run(); err != nil {
		if ee, ok := err.(*exec.ExitError); ok {
			os.Exit(ee.ExitCode())
		}
		fmt.Fprintln(os.Stderr, err)
		os.Exit(1)
	}
}
