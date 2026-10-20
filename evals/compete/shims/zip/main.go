// zip 命令的最小替身：Windows 默认没有 zip，而部分竞品 skill 的脚本会直接调用 `zip -qr 输出.zip .`。
//
// 只实现竞品实际用到的写法：zip [选项] 压缩包路径 输入路径...，选项里只认 -r（递归目录），
// 其余选项（-q、-X、-9 等）一律忽略；压缩包已存在时直接覆盖（调用方在调用前自己删除旧文件）。
// 条目名一律用相对路径和正斜杠，[Content_Types].xml、_rels/.rels 排在最前，便于 Office 类读取器。
// 构建：go build -o zip.exe main.go
package main

import (
	"archive/zip"
	"fmt"
	"io"
	"io/fs"
	"os"
	"path/filepath"
	"sort"
	"strings"
)

func main() {
	var names []string
	for _, a := range os.Args[1:] {
		if strings.HasPrefix(a, "-") && len(a) > 1 {
			continue // 忽略全部选项；递归行为对目录输入始终开启
		}
		names = append(names, a)
	}
	if len(names) < 2 {
		fmt.Fprintln(os.Stderr, "用法：zip [选项] 压缩包 输入路径...")
		os.Exit(16)
	}
	archive, inputs := names[0], names[1:]
	if !strings.HasSuffix(strings.ToLower(archive), ".zip") && filepath.Ext(archive) == "" {
		archive += ".zip"
	}

	// 收集 条目名 -> 磁盘路径
	entries := map[string]string{}
	for _, in := range inputs {
		info, err := os.Stat(in)
		if err != nil {
			fmt.Fprintf(os.Stderr, "zip warning: name not matched: %s\n", in)
			continue
		}
		if !info.IsDir() {
			entries[toEntry(in)] = in
			continue
		}
		filepath.WalkDir(in, func(p string, d fs.DirEntry, err error) error {
			if err != nil || d.IsDir() {
				return nil
			}
			if abs1, e1 := filepath.Abs(p); e1 == nil {
				if abs2, e2 := filepath.Abs(archive); e2 == nil && abs1 == abs2 {
					return nil // 不把压缩包自己打进去
				}
			}
			entries[toEntry(p)] = p
			return nil
		})
	}
	if len(entries) == 0 {
		fmt.Fprintln(os.Stderr, "zip error: Nothing to do!")
		os.Exit(12)
	}
	keys := make([]string, 0, len(entries))
	for k := range entries {
		keys = append(keys, k)
	}
	sort.Slice(keys, func(i, j int) bool {
		ri, rj := rank(keys[i]), rank(keys[j])
		if ri != rj {
			return ri < rj
		}
		return keys[i] < keys[j]
	})

	out, err := os.Create(archive)
	if err != nil {
		fmt.Fprintf(os.Stderr, "zip error: Could not create output file (%v)\n", err)
		os.Exit(15)
	}
	zw := zip.NewWriter(out)
	for _, k := range keys {
		src := entries[k]
		info, err := os.Stat(src)
		if err != nil {
			continue
		}
		h, _ := zip.FileInfoHeader(info)
		h.Name = k
		h.Method = zip.Deflate
		w, err := zw.CreateHeader(h)
		if err != nil {
			fmt.Fprintf(os.Stderr, "zip error: %v\n", err)
			os.Exit(1)
		}
		f, err := os.Open(src)
		if err != nil {
			fmt.Fprintf(os.Stderr, "zip error: %v\n", err)
			os.Exit(1)
		}
		io.Copy(w, f)
		f.Close()
	}
	if err := zw.Close(); err != nil {
		fmt.Fprintf(os.Stderr, "zip error: %v\n", err)
		os.Exit(1)
	}
	out.Close()
}

// toEntry 把输入路径转成压缩包内的条目名：去掉前导 ./，统一正斜杠。
func toEntry(p string) string {
	p = filepath.ToSlash(filepath.Clean(p))
	p = strings.TrimPrefix(p, "./")
	return p
}

// rank 让 OPC 包的关键部件排在最前。
func rank(name string) int {
	switch name {
	case "[Content_Types].xml":
		return 0
	case "_rels/.rels":
		return 1
	}
	return 2
}
