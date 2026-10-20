// unzip 命令的最小替身：Windows 的 PowerShell 里没有 unzip，而 mizzlelover/gongwen-gbt9704-skill 的验证步骤
// 会调用 `unzip -t 文件`、`unzip -p 文件 部件`。只实现这几种写法：
//   unzip -t 压缩包          测试完整性
//   unzip -p 压缩包 部件...   把部件内容写到标准输出
//   unzip -l 压缩包          列出部件
//   unzip [-o] [-q] [-d 目录] 压缩包   解压
// 构建：go build -o unzip.exe main.go
package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"
)

func main() {
	var mode string
	var dir string
	var rest []string
	args := os.Args[1:]
	for i := 0; i < len(args); i++ {
		a := args[i]
		switch {
		case a == "-t" || a == "-p" || a == "-l":
			mode = a
		case a == "-d" && i+1 < len(args):
			dir = args[i+1]
			i++
		case strings.HasPrefix(a, "-") && len(a) > 1:
			// -o -q 等一律忽略；-qo 之类的组合里若含 t/p/l 也认
			if strings.ContainsAny(a, "tpl") && mode == "" {
				for _, c := range a[1:] {
					if c == 't' || c == 'p' || c == 'l' {
						mode = "-" + string(c)
					}
				}
			}
		default:
			rest = append(rest, a)
		}
	}
	if len(rest) < 1 {
		fmt.Fprintln(os.Stderr, "用法：unzip [-t|-p|-l] [-d 目录] 压缩包 [部件...]")
		os.Exit(10)
	}
	zr, err := zip.OpenReader(rest[0])
	if err != nil {
		fmt.Fprintf(os.Stderr, "unzip: cannot find or open %s (%v)\n", rest[0], err)
		os.Exit(9)
	}
	defer zr.Close()
	members := rest[1:]
	want := func(name string) bool {
		if len(members) == 0 {
			return true
		}
		for _, m := range members {
			if m == name {
				return true
			}
		}
		return false
	}
	switch mode {
	case "-t":
		for _, f := range zr.File {
			rc, err := f.Open()
			if err != nil {
				fmt.Printf("    testing: %s   ERROR\n", f.Name)
				os.Exit(2)
			}
			if _, err := io.Copy(io.Discard, rc); err != nil {
				fmt.Printf("    testing: %s   ERROR\n", f.Name)
				os.Exit(2)
			}
			rc.Close()
		}
		fmt.Printf("No errors detected in compressed data of %s.\n", rest[0])
	case "-l":
		fmt.Println("  Length      Name")
		for _, f := range zr.File {
			fmt.Printf("%10d  %s\n", f.UncompressedSize64, f.Name)
		}
	case "-p":
		found := false
		for _, f := range zr.File {
			if want(f.Name) {
				rc, err := f.Open()
				if err != nil {
					os.Exit(2)
				}
				io.Copy(os.Stdout, rc)
				rc.Close()
				found = true
			}
		}
		if !found {
			fmt.Fprintf(os.Stderr, "caution: filename not matched:  %s\n", strings.Join(members, " "))
			os.Exit(11)
		}
	default:
		if dir == "" {
			dir = "."
		}
		for _, f := range zr.File {
			if !want(f.Name) {
				continue
			}
			target := filepath.Join(dir, filepath.FromSlash(f.Name))
			if strings.HasSuffix(f.Name, "/") {
				os.MkdirAll(target, 0o755)
				continue
			}
			os.MkdirAll(filepath.Dir(target), 0o755)
			rc, err := f.Open()
			if err != nil {
				os.Exit(2)
			}
			out, err := os.Create(target)
			if err != nil {
				os.Exit(2)
			}
			io.Copy(out, rc)
			out.Close()
			rc.Close()
		}
	}
}
