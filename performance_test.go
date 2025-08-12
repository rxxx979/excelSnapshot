package main

import (
	"fmt"
	"log"
	"net/http"
	_ "net/http/pprof"
	"os"
	"runtime"
	"runtime/pprof"
	"testing"
	"time"
)

// 性能测试配置
const (
	TestExcelFile = "demo.xlsx"
	TestOutputDir = "."
)

// 启动性能监控服务器
func startProfilingServer() {
	go func() {
		log.Println("性能分析服务启动: http://localhost:6060/debug/pprof/")
		log.Println("使用方法:")
		log.Println("  CPU分析: go tool pprof http://localhost:6060/debug/pprof/profile?seconds=30")
		log.Println("  内存分析: go tool pprof http://localhost:6060/debug/pprof/heap")
		log.Println("  协程分析: go tool pprof http://localhost:6060/debug/pprof/goroutine")
		if err := http.ListenAndServe("localhost:6060", nil); err != nil {
			log.Printf("性能分析服务启动失败: %v", err)
		}
	}()
	time.Sleep(100 * time.Millisecond) // 等待服务启动
}

// 基准测试：单次渲染性能
func BenchmarkSingleRender(b *testing.B) {
	// 确保测试文件存在
	if _, err := os.Stat(TestExcelFile); os.IsNotExist(err) {
		b.Skip("测试Excel文件不存在，跳过性能测试")
	}

	// 创建输出目录
	os.MkdirAll(TestOutputDir, 0755)

	// 初始化字体系统
	initFonts()

	b.ResetTimer()
	b.ReportAllocs()

	for i := 0; i < b.N; i++ {
		outputFile := fmt.Sprintf("%s/bench_%d.png", TestOutputDir, i)
		err := renderPNGFromExcel(TestExcelFile, "", 0, outputFile, false)
		if err != nil {
			b.Fatalf("渲染失败: %v", err)
		}
		// 清理测试文件
		os.Remove(outputFile)
	}
}

// 内存使用测试
func TestMemoryUsage(t *testing.T) {
	if _, err := os.Stat(TestExcelFile); os.IsNotExist(err) {
		t.Skip("测试Excel文件不存在，跳过内存测试")
	}

	// 启动性能监控
	startProfilingServer()

	// 创建输出目录
	os.MkdirAll(TestOutputDir, 0755)
	initFonts()

	// 记录初始内存状态
	var m1, m2 runtime.MemStats
	runtime.GC()
	runtime.ReadMemStats(&m1)

	// 执行多次渲染测试内存使用
	iterations := 10
	for i := 0; i < iterations; i++ {
		outputFile := fmt.Sprintf("%s/memory_test_%d.png", TestOutputDir, i)
		err := renderPNGFromExcel(TestExcelFile, "", 0, outputFile, false)
		if err != nil {
			t.Fatalf("渲染失败: %v", err)
		}
		os.Remove(outputFile) // 立即清理
	}

	// 强制GC并记录最终内存状态
	runtime.GC()
	runtime.ReadMemStats(&m2)

	// 输出内存使用统计
	fmt.Printf("\n=== 内存使用分析 ===\n")
	fmt.Printf("渲染次数: %d\n", iterations)
	fmt.Printf("初始堆内存: %.2f MB\n", float64(m1.HeapAlloc)/1024/1024)
	fmt.Printf("最终堆内存: %.2f MB\n", float64(m2.HeapAlloc)/1024/1024)
	fmt.Printf("内存增长: %.2f MB\n", float64(m2.HeapAlloc-m1.HeapAlloc)/1024/1024)
	fmt.Printf("总分配内存: %.2f MB\n", float64(m2.TotalAlloc-m1.TotalAlloc)/1024/1024)
	fmt.Printf("GC次数: %d\n", m2.NumGC-m1.NumGC)
	fmt.Printf("平均每次渲染分配: %.2f MB\n", float64(m2.TotalAlloc-m1.TotalAlloc)/1024/1024/float64(iterations))
}

// CPU使用分析测试
func TestCPUProfile(t *testing.T) {
	if _, err := os.Stat(TestExcelFile); os.IsNotExist(err) {
		t.Skip("测试Excel文件不存在，跳过CPU分析")
	}

	// 创建CPU profile文件
	cpuFile, err := os.Create("cpu_profile.prof")
	if err != nil {
		t.Fatalf("无法创建CPU profile文件: %v", err)
	}
	defer cpuFile.Close()

	// 开始CPU profiling
	if err := pprof.StartCPUProfile(cpuFile); err != nil {
		t.Fatalf("无法启动CPU profiling: %v", err)
	}
	defer pprof.StopCPUProfile()

	// 创建输出目录
	os.MkdirAll(TestOutputDir, 0755)
	initFonts()

	// 执行渲染任务
	start := time.Now()
	iterations := 5
	for i := 0; i < iterations; i++ {
		outputFile := fmt.Sprintf("%s/cpu_test_%d.png", TestOutputDir, i)
		err := renderPNGFromExcel(TestExcelFile, "", 0, outputFile, false)
		if err != nil {
			t.Fatalf("渲染失败: %v", err)
		}
		os.Remove(outputFile)
	}
	duration := time.Since(start)

	fmt.Printf("\n=== CPU性能分析 ===\n")
	fmt.Printf("渲染次数: %d\n", iterations)
	fmt.Printf("总耗时: %v\n", duration)
	fmt.Printf("平均每次: %v\n", duration/time.Duration(iterations))
	fmt.Printf("CPU profile已保存到: cpu_profile.prof\n")
	fmt.Printf("查看方法: go tool pprof cpu_profile.prof\n")
}

// 并发性能测试
func TestConcurrentPerformance(t *testing.T) {
	if _, err := os.Stat(TestExcelFile); os.IsNotExist(err) {
		t.Skip("测试Excel文件不存在，跳过并发测试")
	}

	os.MkdirAll(TestOutputDir, 0755)
	initFonts()

	// 测试不同并发度的性能
	concurrencies := []int{1, 2, 4, 8}
	iterations := 8

	for _, concurrency := range concurrencies {
		t.Run(fmt.Sprintf("Concurrency_%d", concurrency), func(t *testing.T) {
			start := time.Now()

			// 使用信号量控制并发度
			semaphore := make(chan struct{}, concurrency)
			done := make(chan bool, iterations)

			for i := 0; i < iterations; i++ {
				go func(index int) {
					semaphore <- struct{}{}        // 获取信号量
					defer func() { <-semaphore }() // 释放信号量

					outputFile := fmt.Sprintf("%s/concurrent_%d_%d.png", TestOutputDir, concurrency, index)
					err := renderPNGFromExcel(TestExcelFile, "", 0, outputFile, false)
					if err != nil {
						t.Errorf("渲染失败: %v", err)
					}
					os.Remove(outputFile)
					done <- true
				}(i)
			}

			// 等待所有任务完成
			for i := 0; i < iterations; i++ {
				<-done
			}

			duration := time.Since(start)
			fmt.Printf("并发度 %d: 总耗时 %v, 平均 %v\n",
				concurrency, duration, duration/time.Duration(iterations))
		})
	}
}

// 内存泄漏检测
func TestMemoryLeak(t *testing.T) {
	if _, err := os.Stat(TestExcelFile); os.IsNotExist(err) {
		t.Skip("测试Excel文件不存在，跳过内存泄漏测试")
	}

	os.MkdirAll(TestOutputDir, 0755)
	initFonts()

	// 记录每轮的内存使用
	rounds := 5
	iterations := 10
	memoryUsage := make([]uint64, rounds)

	for round := 0; round < rounds; round++ {
		// 执行多次渲染
		for i := 0; i < iterations; i++ {
			outputFile := fmt.Sprintf("%s/leak_test_%d_%d.png", TestOutputDir, round, i)
			err := renderPNGFromExcel(TestExcelFile, "", 0, outputFile, false)
			if err != nil {
				t.Fatalf("渲染失败: %v", err)
			}
			os.Remove(outputFile)
		}

		// 强制GC并记录内存使用
		runtime.GC()
		runtime.GC() // 两次GC确保清理完毕
		var m runtime.MemStats
		runtime.ReadMemStats(&m)
		memoryUsage[round] = m.HeapAlloc
	}

	// 分析内存使用趋势
	fmt.Printf("\n=== 内存泄漏检测 ===\n")
	for i, usage := range memoryUsage {
		fmt.Printf("第%d轮后内存使用: %.2f MB\n", i+1, float64(usage)/1024/1024)
	}

	// 检查内存是否持续增长
	if memoryUsage[rounds-1] > memoryUsage[0]*2 {
		t.Logf("警告: 内存使用增长较大，可能存在内存泄漏")
	} else {
		t.Logf("内存使用稳定，未发现明显泄漏")
	}
}

// 清理测试文件
func TestCleanup(t *testing.T) {
	os.RemoveAll(TestOutputDir)
	os.Remove("cpu_profile.prof")
	fmt.Println("测试文件清理完成")
}
